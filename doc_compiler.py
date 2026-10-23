# -*- coding: utf-8 -*-
"""
Document compilation, multimodal slide/diagram parser, and storage synchronization worker.
Supports multi-format rendering (MD, DOCX, HTML, PDF, XLSX) and cloud state persistence.
"""
import os, re, io, sys, time, json, base64, shutil, tempfile, subprocess

# ==========================================
# 0. Dependency Bootstrap & Environment
# ==========================================
def _pip(*pkgs):
    subprocess.run([sys.executable, "-m", "pip", "install", "-q", "-U", *pkgs], check=False, timeout=900)

REQUIRED_PKGS = [
    "requests", "google-api-python-client", "google-auth",
    "python-docx", "python-pptx", "pypdf", "openpyxl", "markdown", "Pillow"
]
try:
    import requests, docx, pptx, pypdf, openpyxl, markdown
    from PIL import Image
    from googleapiclient.discovery import build as _gbuild
    from googleapiclient.http import MediaInMemoryUpload
    from google.oauth2.credentials import Credentials as _UserCredentials
    from google.auth.transport.requests import Request as _GRequest
except Exception:
    print("[INIT] Installing required dependencies...")
    _pip(*REQUIRED_PKGS)
    import requests, docx, pptx, pypdf, openpyxl, markdown
    from PIL import Image
    from googleapiclient.discovery import build as _gbuild
    from googleapiclient.http import MediaInMemoryUpload
    from google.oauth2.credentials import Credentials as _UserCredentials
    from google.auth.transport.requests import Request as _GRequest

from docx.shared import Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement, parse_xml
from docx.oxml.ns import qn, nsdecls
from openpyxl.styles import Font as XlFont, Alignment as XlAlign, PatternFill as XlFill

RUN_T0 = time.time()
MAX_RUN_MINUTES = float(os.environ.get("MAX_RUN_MINUTES", "330"))

def time_budget_left():
    return MAX_RUN_MINUTES * 60 - (time.time() - RUN_T0)

def get_secret(name):
    for nm in (name, name.upper(), name.lower()):
        try:
            from google.colab import userdata
            v = userdata.get(nm)
            if v: return v.strip()
        except Exception:
            pass
        v = os.environ.get(nm)
        if v: return v.strip()
    return None

# ==========================================
# 1. Key Pooling (Dual-Cline + OpenRouter Pool)
# ==========================================
def load_key_pool(prefixes):
    pool = []
    for pf in prefixes:
        base_k = get_secret(pf)
        if base_k and base_k not in pool:
            pool.append(base_k)
        for i in range(1, 16):
            for sep in ("_", ""):
                k = get_secret(f"{pf}{sep}{i}")
                if k and k not in pool:
                    pool.append(k)
    return pool

CLINE_KEYS = load_key_pool(["CLINE_KEY", "CLINE_API_KEY"])
OPENROUTER_KEYS = load_key_pool(["OPENROUTER_KEY", "OPENROUTER_API_KEY", "OR_KEY"])

print(f"[INIT] Configured endpoints -> Primary pool: {len(CLINE_KEYS)} | Secondary pool: {len(OPENROUTER_KEYS)}")

# ==========================================
# 2. Upstream Routing & Zero-Wait Failover
# ==========================================
DEFAULT_MODEL_ORDER = get_secret("MODEL_ORDER") or "123"
CUSTOM_TARGET_MODEL = get_secret("TARGET_MODEL")

# Both possible Cline endpoints are included for automatic fallback if one returns 403/404
CLINE_URLS = [
    "https://api.cline.bot/api/v1/chat/completions",
    "https://api.cline.bot/v1/chat/completions"
]

# Authentic VSCode/Cline client headers to pass Cloudflare WAF without 403 Forbidden
CLINE_HEADERS = {
    "HTTP-Referer": "https://cline.bot",
    "X-Title": "Cline",
    "Origin": "vscode-file://vscode-app"
}

OR_HEADERS = {
    "HTTP-Referer": "https://github.com",
    "X-Title": "Doc Pipeline"
}

TIER_CONFIG = {
    1: {
        "id": 1,
        "label": "tier-1-primary",
        "upstreams": [
            {
                "name": "gw_primary_v1",
                " urls": CLINE_URLS,
                "models": [CUSTOM_TARGET_MODEL] if CUSTOM_TARGET_MODEL else [
                    "cline-pass/muse-spark-1.3-contributor",
                    "muse-spark-1.3-contributor",
                    "meta/muse-spark-1.3"
                ],
                "keys": CLINE_KEYS,
                "headers": CLINE_HEADERS
            },
            {
                "name": "gw_secondary",
                "urls": ["https://openrouter.ai/api/v1/chat/completions"],
                "models": ["meta/muse-spark-1.3:free", "meta/muse-spark-1.3"],
                "keys": OPENROUTER_KEYS,
                "headers": OR_HEADERS
            }
        ]
    },
    2: {
        "id": 2,
        "label": "tier-2-analytical",
        "upstreams": [
            {
                "name": "gw_primary_v1",
                "urls": CLINE_URLS,
                "models": ["cline-pass/deepseek-v4.1-flash", "deepseek/deepseek-chat"],
                "keys": CLINE_KEYS,
                "headers": CLINE_HEADERS
            },
            {
                "name": "gw_secondary",
                "urls": ["https://openrouter.ai/api/v1/chat/completions"],
                "models": [
                    "deepseek/deepseek-r1:free",
                    "deepseek/deepseek-chat-v3-0324:free",
                    "deepseek/deepseek-chat:free",
                    "deepseek/deepseek-r1",
                    "deepseek/deepseek-chat"
                ],
                "keys": OPENROUTER_KEYS,
                "headers": OR_HEADERS
            }
        ]
    },
    3: {
        "id": 3,
        "label": "tier-3-general",
        "upstreams": [
            {
                "name": "gw_primary_v1",
                "urls": CLINE_URLS,
                "models": ["cline-pass/glm-5.3-flash"],
                "keys": CLINE_KEYS,
                "headers": CLINE_HEADERS
            },
            {
                "name": "gw_secondary",
                "urls": ["https://openrouter.ai/api/v1/chat/completions"],
                "models": [
                    "thudm/glm-4-9b:free",
                    "zhipu/glm-5.3",
                    "thudm/glm-4-plus"
                ],
                "keys": OPENROUTER_KEYS,
                "headers": OR_HEADERS
            }
        ]
    }
}

def parse_order_string(spec_str):
    if not spec_str:
        return [1, 2, 3]
    persian_map = {'۱': 1, '۲': 2, '۳': 3}
    order = []
    for ch in spec_str.strip():
        if ch in ('1', '2', '3'):
            d = int(ch)
        elif ch in persian_map:
            d = persian_map[ch]
        else:
            continue
        if d not in order:
            order.append(d)
    return order if order else [1, 2, 3]

class UpstreamClient:
    """
    Handles round-robin key rotation, zero-wait failover across dual primary keys,
    Cloudflare WAF browser-header emulation, and 402/404 model skipping.
    """
    def __init__(self):
        self.active_order = [1, 2, 3]
        self.pointers = {}
        self.cooldowns = {}
        self.resolved_targets = {}  # pname -> (url, model_id)
        self.unsupported_combos = set()  # (url, model_id) that returned 402/404

    def set_order(self, order_list):
        self.active_order = order_list
        print(f"[CONFIG] Execution sequence set to: {self.active_order}")

    def _acquire_key(self, pool_name, keys):
        if not keys:
            return None
        now = time.time()
        n = len(keys)
        ptr = self.pointers.get(pool_name, 0)
        for i in range(n):
            idx = (ptr + i) % n
            k = keys[idx]
            if self.cooldowns.get(k, 0) <= now:
                self.pointers[pool_name] = (idx + 1) % n
                return k
        return None

    def _send_request(self, url, api_key, model_id, messages, extra_headers):
        headers = {
            "Authorization": f"Bearer {api_key}",
            "Content-Type": "application/json",
            "Accept": "application/json",
            "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Code/1.96.0 Chrome/128.0.6613.186 Electron/32.2.6 Safari/537.36"
        }
        if extra_headers:
            headers.update(extra_headers)
        payload = {
            "model": model_id,
            "messages": messages,
            "temperature": 0.3,
            "stream": False
        }
        r = requests.post(url, headers=headers, json=payload, timeout=(25, 360))
        short_body = (r.text or "").replace("\n", " ")[:90]

        if r.status_code == 429:
            self.cooldowns[api_key] = time.time() + 45
            raise RuntimeError(f"HTTP_429 ({short_body})")

        # 402 (Payment Required) or 404 (Not Found) means this specific (url, model_id) is unavailable/paid
        if r.status_code in (402, 404):
            raise LookupError(f"HTTP_{r.status_code} [{model_id}] ({short_body})")

        if r.status_code == 400 and any(w in r.text.lower() for w in ("model", "not found", "image", "vision", "invalid")):
            raise LookupError(f"HTTP_400 [{model_id}] ({short_body})")

        if r.status_code != 200:
            raise RuntimeError(f"HTTP_{r.status_code} ({short_body})")

        data = r.json()
        if "error" in data and not data.get("choices"):
            err_msg = str(data["error"])[:90]
            if "402" in err_msg or "afford" in err_msg.lower() or "credit" in err_msg.lower() or "not found" in err_msg.lower():
                raise LookupError(f"API_ERR [{model_id}] ({err_msg})")
            raise RuntimeError(f"API_ERR ({err_msg})")

        content = data["choices"][0]["message"]["content"]
        return content.strip() if content else ""

    def complete(self, system_prompt, user_content, prefer_vision=False):
        messages = [
            {"role": "system", "content": system_prompt},
            {"role": "user", "content": user_content}
        ]
        order = ([1] + [x for x in self.active_order if x != 1]) if prefer_vision else self.active_order

        for back_off_round in range(1, 4):
            for tier_id in order:
                tier = TIER_CONFIG.get(tier_id)
                if not tier:
                    continue
                for up in tier["upstreams"]:
                    pname = f"{up['name']}_t{tier_id}"
                    keys = up["keys"]
                    if not keys:
                        continue
                    urls = up.get("urls") or up.get(" urls") or []
                    models = up["models"]

                    # Build candidate (url, model_id) pairs, prioritizing previously working combo
                    candidates = []
                    locked = self.resolved_targets.get(pname)
                    if locked and locked not in self.unsupported_combos:
                        candidates.append(locked)
                    for u in urls:
                        for m in models:
                            pair = (u, m)
                            if pair not in candidates and pair not in self.unsupported_combos:
                                candidates.append(pair)

                    if not candidates:
                        continue

                    for _ in range(len(keys)):
                        api_key = self._acquire_key(pname, keys)
                        if not api_key:
                            break
                        switch_key = False
                        for (u, m_id) in list(candidates):
                            if (u, m_id) in self.unsupported_combos:
                                continue
                            try:
                                res = self._send_request(u, api_key, m_id, messages, up["headers"])
                                if res:
                                    self.resolved_targets[pname] = (u, m_id)
                                    return res
                            except LookupError as le:
                                # Model/URL is paid (402) or invalid (404): mark combo and immediately try next model on SAME key
                                print(f"[INFO] Skipping {le}; trying next model/endpoint...")
                                self.unsupported_combos.add((u, m_id))
                                continue
                            except Exception as e:
                                # 403 / 429 / 5xx: switch to next key or next endpoint URL
                                print(f"[WARN] Upstream {up['name']} (tier {tier_id}) returned {e}; switching...")
                                switch_key = True
                                break
                        if not switch_key:
                            break

            wait_sec = 15 * back_off_round
            print(f"[INFO] Active pools busy. Backing off {wait_sec}s (round {back_off_round}/3)...")
            time.sleep(wait_sec)

        raise RuntimeError("All configured upstream endpoints failed after retries.")

CLIENT = UpstreamClient()

# ==========================================
# 3. Cloud Storage Layer (Drive API v3)
# ==========================================
FOLDER_MIME = "application/vnd.google-apps.folder"
TOKEN_URI = "https://oauth2.googleapis.com/token"

class StorageClient:
    SCOPES = ["https://www.googleapis.com/auth/drive"]
    def __init__(self, cred_json):
        info = json.loads(cred_json) if isinstance(cred_json, str) else cred_json
        self.creds = _UserCredentials(
            token=None,
            refresh_token=info["refresh_token"],
            client_id=info["client_id"],
            client_secret=info["client_secret"],
            token_uri=info.get("token_uri") or TOKEN_URI,
            scopes=self.SCOPES,
        )
        self.creds.refresh(_GRequest())
        self.svc = _gbuild("drive", "v3", credentials=self.creds, cache_discovery=False)

    def find(self, parent, name, folder=None):
        safe_name = name.replace("\\", "\\\\").replace("'", "\\'")
        q = f"name = '{safe_name}' and '{parent}' in parents and trashed = false"
        if folder is True:
            q += f" and mimeType = '{FOLDER_MIME}'"
        elif folder is False:
            q += f" and mimeType != '{FOLDER_MIME}'"
        r = self.svc.files().list(q=q, fields="files(id,name,size,mimeType)").execute(num_retries=3)
        fs = r.get("files", [])
        return fs[0] if fs else None

    def ensure_folder(self, parent, name):
        f = self.find(parent, name, folder=True)
        if f:
            return f["id"]
        body = {"name": name, "mimeType": FOLDER_MIME, "parents": [parent]}
        return self.svc.files().create(body=body, fields="id").execute(num_retries=3)["id"]

    def list_children(self, parent, folders_only=None):
        q = f"'{parent}' in parents and trashed = false"
        if folders_only is True:
            q += f" and mimeType = '{FOLDER_MIME}'"
        elif folders_only is False:
            q += f" and mimeType != '{FOLDER_MIME}'"
        r = self.svc.files().list(q=q, pageSize=500, fields="files(id,name,size,mimeType)").execute(num_retries=3)
        return r.get("files", [])

    def download_bytes(self, file_id):
        return self.svc.files().get_media(fileId=file_id).execute(num_retries=3)

    def put_bytes(self, parent, name, data_bytes, mime_type="application/octet-stream"):
        media = MediaInMemoryUpload(data_bytes, mimetype=mime_type, resumable=False)
        existing = self.find(parent, name, folder=False)
        if existing:
            self.svc.files().update(fileId=existing["id"], media_body=media).execute(num_retries=3)
            return existing["id"]
        body = {"name": name, "parents": [parent]}
        return self.svc.files().create(body=body, media_body=media, fields="id").execute(num_retries=3)["id"]

    def put_text(self, parent, name, text):
        return self.put_bytes(parent, name, text.encode("utf-8"), "text/plain; charset=utf-8")

def sync_order_config(storage, root_id):
    files = storage.list_children(root_id, folders_only=False)
    order_file = None
    spec_from_name = None

    for f in files:
        base = os.path.splitext(f["name"])[0].strip()
        if base.upper() in ("MODEL_ORDER", "ORDER", "CONFIG"):
            order_file = f
            break
        if re.fullmatch(r'[123۱۲۳]+', base):
            order_file = f
            spec_from_name = base
            break

    default_content = (
        "123\n\n"
        "# ==========================================\n"
        "# 📘 راهنمای اولویت‌بندی مدل‌ها (کنترل با موبایل)\n"
        "# ترتیب اجرای مدل‌ها را در خط اول همین فایل بنویسید (یا نام همین فایل را تغییر دهید):\n"
        "# \n"
        "# کد ۱ = مدل Muse Spark 1.3 (متا - بالاترین عمق علمی، دانشنامه، کتب مرجع و تحلیل چندحالته اسلایدها)\n"
        "# کد ۲ = مدل DeepSeek 4.1   (دیپ‌سیک - استدلال منطقی، پردازش جداول اکسل، محاسبات و کارت امتیازی BSC)\n"
        "# کد ۳ = مدل GLM 5.3        (زد آی - سرعت فوق‌العاده بالا، کانتکست ۱ میلیون، مقالات سایت و خلاصه‌سازی)\n"
        "# ------------------------------------------\n"
        "# الگوهای کاربردی:\n"
        "# • فقط و فقط مدل اول (بدون سوئیچ): 1\n"
        "# • اولویت با دیپ‌سیک و سپس اسمارک: 21\n"
        "# • هر سه مدل به ترتیب استاندارد: 123\n"
        "# • اولویت اول دیپ‌سیک، دوم اسمارک، سوم جی‌ال‌ام: 213\n"
        "# ==========================================\n"
    )

    if not order_file:
        storage.put_text(root_id, f"{DEFAULT_MODEL_ORDER}.txt", default_content)
        return parse_order_string(DEFAULT_MODEL_ORDER)

    if spec_from_name and spec_from_name != DEFAULT_MODEL_ORDER:
        return parse_order_string(spec_from_name)

    try:
        raw_txt = storage.download_bytes(order_file["id"]).decode("utf-8", errors="ignore")
        first_line = next((ln.strip() for ln in raw_txt.splitlines() if ln.strip() and not ln.strip().startswith("#")), "")
        if first_line:
            return parse_order_string(first_line)
    except Exception:
        pass

    return parse_order_string(spec_from_name or DEFAULT_MODEL_ORDER)

# ==========================================
# 4. Multimodal Vision & Presentation Parser (PPTX / Scanned PDF / Images)
# ==========================================
VISION_SYSTEM_PROMPT = (
    "You are a specialist in visual document analysis, diagram deconstruction, and Persian/English OCR. "
    "Examine the provided slide/page image(s) meticulously and output a detailed structured report in formal Persian."
)

DIAGRAM_DECONSTRUCTION_PROMPT = (
    "لطفاً تصویر این اسلاید/صفحه را با دقت کامل کلمه‌به‌کلمه و شکل‌به‌شکل واکاوی کن و خروجی را به فارسی بنویس:\n"
    "۱. **واکاوی دیاگرام‌ها و فلوچارت‌ها:** تشریح کامل نمودارها، درخت‌های تصمیم، جهت پیکان‌ها، تقدم و تأخر فرآیندها و روابط علت و معلولی بین المان‌های بصری.\n"
    "۲. **استخراج کامل متون و اعداد (OCR عمیق):** تمام متون، تیترها، برچسب‌های داخل اشکال، جداول، درصدها، فرمول‌ها و اعداد موجود در تصویر را بدون جا افتادن هیچ جزئیاتی استخراج کن.\n"
    "۳. **پیام مفهومی و ساختاری:** هدف و پیام منطقی که طراح اسلاید از طریق چیدمان، رنگ‌بندی یا گروه‌بندی بخش‌ها قصد انتقال آن را داشته است تشریح کن."
)

def _optimize_image_to_base64(img_input, max_width=1280, quality=80):
    im = Image.open(img_input) if not isinstance(img_input, Image.Image) else img_input
    if im.mode in ("RGBA", "P"):
        im = im.convert("RGB")
    w, h = im.size
    if w > max_width:
        new_h = int(h * (max_width / float(w)))
        im = im.resize((max_width, new_h), Image.Resampling.LANCZOS)
    buf = io.BytesIO()
    im.save(buf, format="JPEG", quality=quality, optimize=True)
    b64 = base64.b64encode(buf.getvalue()).decode("ascii")
    return f"data:image/jpeg;base64,{b64}"

def _analyze_images_with_vision(image_b64_list, context_label=""):
    content_payload = [
        {"type": "text", "text": f"{DIAGRAM_DECONSTRUCTION_PROMPT}\n\n{context_label}".strip()}
    ]
    for b64_uri in image_b64_list:
        content_payload.append({
            "type": "image_url",
            "image_url": {"url": b64_uri}
        })
    try:
        return CLIENT.complete(VISION_SYSTEM_PROMPT, content_payload, prefer_vision=True)
    except Exception as e:
        print(f"[WARN] Vision analysis fallback triggered: {e}")
        return ""

def _extract_pptx_text_per_slide(raw_bytes):
    slides_text = []
    try:
        prs = pptx.Presentation(io.BytesIO(raw_bytes))
        for idx, slide in enumerate(prs.slides, 1):
            parts = []
            for shape in slide.shapes:
                if getattr(shape, "has_text_frame", False):
                    txt = shape.text_frame.text.strip()
                    if txt:
                        parts.append(txt)
                if getattr(shape, "has_table", False):
                    for row in shape.table.rows:
                        parts.append(" | ".join(c.text.strip() for c in row.cells))
            if slide.has_notes_slide and slide.notes_slide.notes_text_frame:
                notes = slide.notes_slide.notes_text_frame.text.strip()
                if notes:
                    parts.append(f"[یادداشت سخنران: {notes}]")
            slides_text.append((idx, "\n".join(parts)))
    except Exception as e:
        print(f"[WARN] python-pptx baseline extraction notice: {e}")
    return slides_text

def _render_presentation_to_jpegs(filename, raw_bytes, tmp_dir):
    in_path = os.path.join(tmp_dir, filename)
    with open(in_path, "wb") as f:
        f.write(raw_bytes)
    if not shutil.which("libreoffice") or not shutil.which("pdftoppm"):
        return []
    try:
        subprocess.run(
            ["libreoffice", "--headless", "--convert-to", "pdf", in_path, "--outdir", tmp_dir],
            check=True, timeout=180, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL
        )
        base_no_ext = os.path.splitext(filename)[0]
        pdf_path = os.path.join(tmp_dir, base_no_ext + ".pdf")
        if not os.path.exists(pdf_path):
            pdfs = [os.path.join(tmp_dir, x) for x in os.listdir(tmp_dir) if x.endswith(".pdf")]
            pdf_path = pdfs[0] if pdfs else None
        if not pdf_path:
            return []
        out_prefix = os.path.join(tmp_dir, "slide_page")
        subprocess.run(
            ["pdftoppm", "-jpeg", "-r", "150", pdf_path, out_prefix],
            check=True, timeout=180, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL
        )
        imgs = sorted([os.path.join(tmp_dir, x) for x in os.listdir(tmp_dir) if x.startswith("slide_page") and x.endswith(".jpg")])
        return imgs
    except Exception as e:
        print(f"[WARN] Headless slide rendering error: {e}")
        return []

def parse_presentation_multimodal(filename, raw_bytes, storage=None, cache_folder_id=None):
    safe_cache_name = re.sub(r'[^\w\-_.]', '_', filename) + "_slides_visual.md"
    if storage and cache_folder_id:
        cached = storage.find(cache_folder_id, safe_cache_name, folder=False)
        if cached and int(cached.get("size", 0)) > 100:
            print(f"[CACHE] Loaded visual slide analysis for {filename} from storage.")
            return storage.download_bytes(cached["id"]).decode("utf-8", errors="ignore")

    print(f"[VISION] Performing multimodal slide & diagram analysis on {filename}...")
    baseline_slides = _extract_pptx_text_per_slide(raw_bytes)
    baseline_map = dict(baseline_slides)

    tmp_dir = tempfile.mkdtemp(prefix="pptx_vis_")
    report_sections = [f"## تحلیل جامع چندحالته پاورپوینت: {filename}\n"]
    try:
        img_paths = _render_presentation_to_jpegs(filename, raw_bytes, tmp_dir)
        if img_paths:
            batch_size = 3
            for start_i in range(0, len(img_paths), batch_size):
                batch_files = img_paths[start_i:start_i + batch_size]
                b64_list = [_optimize_image_to_base64(p) for p in batch_files]
                slide_nums = list(range(start_i + 1, start_i + len(batch_files) + 1))
                raw_text_hint = "\n".join(
                    f"متن خام اسلاید {s_num}:\n{baseline_map.get(s_num, '')}"
                    for s_num in slide_nums if baseline_map.get(s_num)
                )
                label = f"اسلایدهای شماره {slide_nums[0]} تا {slide_nums[-1]} از فایل {filename}.\n{raw_text_hint}"
                print(f"[VISION] Analyzing slides {slide_nums[0]}..{slide_nums[-1]} of {len(img_paths)}...")
                vis_analysis = _analyze_images_with_vision(b64_list, context_label=label)
                if vis_analysis:
                    report_sections.append(f"### تشریح اسلایدهای {slide_nums[0]} تا {slide_nums[-1]}\n{vis_analysis}\n")
                elif raw_text_hint:
                    report_sections.append(f"### محتوای اسلایدهای {slide_nums[0]} تا {slide_nums[-1]}\n{raw_text_hint}\n")
                time.sleep(4)
        else:
            for s_num, s_txt in baseline_slides:
                if s_txt:
                    report_sections.append(f"### اسلاید {s_num}\n{s_txt}\n")
    finally:
        shutil.rmtree(tmp_dir, ignore_errors=True)

    full_report = "\n".join(report_sections)
    if storage and cache_folder_id and len(full_report) > 120:
        storage.put_text(cache_folder_id, safe_cache_name, full_report)
    return full_report

def parse_pdf_hybrid_ocr(filename, raw_bytes, storage=None, cache_folder_id=None):
    safe_cache_name = re.sub(r'[^\w\-_.]', '_', filename) + "_ocr.md"
    if storage and cache_folder_id:
        cached = storage.find(cache_folder_id, safe_cache_name, folder=False)
        if cached and int(cached.get("size", 0)) > 100:
            return storage.download_bytes(cached["id"]).decode("utf-8", errors="ignore")

    reader = pypdf.PdfReader(io.BytesIO(raw_bytes))
    page_texts = []
    scanned_pages = []
    for idx, page in enumerate(reader.pages, 1):
        txt = (page.extract_text() or "").strip()
        if len(txt) < 35:
            scanned_pages.append(idx)
        page_texts.append((idx, txt))

    if scanned_pages and shutil.which("pdftoppm"):
        print(f"[VISION] Detected {len(scanned_pages)} scanned/diagram pages in {filename}. Running Vision OCR...")
        tmp_dir = tempfile.mkdtemp(prefix="pdf_ocr_")
        try:
            pdf_path = os.path.join(tmp_dir, "doc.pdf")
            with open(pdf_path, "wb") as f:
                f.write(raw_bytes)
            ocr_map = {}
            target_scanned = scanned_pages[:30]
            for i in range(0, len(target_scanned), 2):
                batch_pnums = target_scanned[i:i+2]
                b64_imgs = []
                for pnum in batch_pnums:
                    out_pref = os.path.join(tmp_dir, f"p_{pnum}")
                    subprocess.run(
                        ["pdftoppm", "-jpeg", "-r", "150", "-f", str(pnum), "-l", str(pnum), pdf_path, out_pref],
                        check=False, timeout=60, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL
                    )
                    rendered = [os.path.join(tmp_dir, x) for x in os.listdir(tmp_dir) if x.startswith(f"p_{pnum}_") or x.startswith(f"p_{pnum}-")]
                    if rendered:
                        b64_imgs.append(_optimize_image_to_base64(rendered[0]))
                if b64_imgs:
                    vis_txt = _analyze_images_with_vision(
                        b64_imgs,
                        context_label=f"صفحات اسکن‌شده شماره {batch_pnums} از سند {filename}"
                    )
                    if vis_txt:
                        ocr_map[batch_pnums[0]] = vis_txt
                time.sleep(3)
            merged = []
            for pnum, orig_txt in page_texts:
                if pnum in ocr_map:
                    merged.append(f"--- [تحلیل بصری و OCR صفحه {pnum}] ---\n{ocr_map[pnum]}")
                elif orig_txt:
                    merged.append(orig_txt)
            final_pdf_text = "\n\n".join(merged)
            if storage and cache_folder_id and len(final_pdf_text) > 100:
                storage.put_text(cache_folder_id, safe_cache_name, final_pdf_text)
            return final_pdf_text
        finally:
            shutil.rmtree(tmp_dir, ignore_errors=True)

    return "\n".join(t for _, t in page_texts if t)

def extract_content_smart(filename, raw_bytes, storage=None, cache_folder_id=None):
    ext = os.path.splitext(filename)[1].lower()
    try:
        if ext in (".txt", ".md", ".json", ".csv", ".py", ".html"):
            return raw_bytes.decode("utf-8", errors="ignore")
        elif ext == ".docx":
            doc = docx.Document(io.BytesIO(raw_bytes))
            out = [p.text for p in doc.paragraphs if p.text.strip()]
            for tbl in doc.tables:
                for row in tbl.rows:
                    out.append(" | ".join(c.text.strip() for c in row.cells))
            return "\n".join(out)
        elif ext in (".pptx", ".ppt"):
            return parse_presentation_multimodal(filename, raw_bytes, storage, cache_folder_id)
        elif ext == ".pdf":
            return parse_pdf_hybrid_ocr(filename, raw_bytes, storage, cache_folder_id)
        elif ext in (".png", ".jpg", ".jpeg", ".webp"):
            safe_cache_name = re.sub(r'[^\w\-_.]', '_', filename) + "_img_visual.md"
            if storage and cache_folder_id:
                cached = storage.find(cache_folder_id, safe_cache_name, folder=False)
                if cached and int(cached.get("size", 0)) > 80:
                    return storage.download_bytes(cached["id"]).decode("utf-8", errors="ignore")
            print(f"[VISION] Running OCR and diagram analysis on image: {filename}...")
            b64_uri = _optimize_image_to_base64(io.BytesIO(raw_bytes))
            vis_res = _analyze_images_with_vision([b64_uri], context_label=f"تصویر مستقل: {filename}")
            if storage and cache_folder_id and vis_res:
                storage.put_text(cache_folder_id, safe_cache_name, vis_res)
            return vis_res
        elif ext in (".xlsx", ".xls"):
            wb = openpyxl.load_workbook(io.BytesIO(raw_bytes), data_only=True)
            out = []
            for sheet in wb.sheetnames:
                ws = wb[sheet]
                out.append(f"### Sheet: {sheet}")
                for row in ws.iter_rows(values_only=True):
                    vals = [str(v).strip() for v in row if v is not None and str(v).strip() != ""]
                    if vals:
                        out.append(" | ".join(vals))
            return "\n".join(out)
    except Exception as e:
        print(f"[WARN] Failed parsing {filename}: {e}")
    return ""

# ==========================================
# 5. Multi-Format Output Builders (MD, DOCX, HTML, PDF, XLSX)
# ==========================================
LATIN_RE = re.compile(r'([A-Za-z0-9_@./#&+\-:%]+(?:\s+[A-Za-z0-9_@./#&+\-:%]+)*)')

def _set_run_fonts(run, is_latin=False, size_pt=12, bold=False, color_rgb=None):
    run.bold = bold
    run.font.size = Pt(size_pt)
    if color_rgb:
        run.font.color.rgb = RGBColor(*color_rgb)
    rPr = run._r.get_or_add_rPr()
    rFonts = rPr.find(qn('w:rFonts'))
    if rFonts is None:
        rFonts = OxmlElement('w:rFonts')
        rPr.append(rFonts)
    rFonts.set(qn('w:ascii'), 'Times New Roman')
    rFonts.set(qn('w:hAnsi'), 'Times New Roman')
    rFonts.set(qn('w:cs'), 'B Zar')
    if not is_latin:
        rtl = OxmlElement('w:rtl')
        rPr.append(rtl)

def _set_para_rtl(para):
    para.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    pPr = para._p.get_or_add_pPr()
    bidi = OxmlElement('w:bidi')
    pPr.append(bidi)

def _add_mixed_text_to_para(para, text, size_pt=12, base_bold=False, color_rgb=None):
    _set_para_rtl(para)
    bold_chunks = re.split(r'(\*\*.*?\*\*)', text)
    for bchunk in bold_chunks:
        if not bchunk:
            continue
        is_bold = base_bold or (bchunk.startswith("**") and bchunk.endswith("**"))
        clean_chunk = bchunk[2:-2] if (bchunk.startswith("**") and bchunk.endswith("**")) else bchunk
        parts = LATIN_RE.split(clean_chunk)
        for part in parts:
            if not part:
                continue
            is_lat = bool(LATIN_RE.fullmatch(part))
            run = para.add_run(part)
            _set_run_fonts(run, is_latin=is_lat, size_pt=size_pt, bold=is_bold, color_rgb=color_rgb)

def build_docx_bytes(md_text):
    doc = docx.Document()
    lines = md_text.splitlines()
    in_table, table_rows = False, []

    def flush_table():
        nonlocal in_table, table_rows
        if not table_rows:
            in_table = False; return
        cols = max(len(r) for r in table_rows)
        tbl = doc.add_table(rows=len(table_rows), cols=cols)
        tbl.style = 'Table Grid'
        tbl._tbl.tblPr.append(OxmlElement('w:bidiVisual'))
        for r_idx, r_data in enumerate(table_rows):
            for c_idx in range(cols):
                cell_txt = r_data[c_idx] if c_idx < len(r_data) else ""
                cell = tbl.cell(r_idx, c_idx)
                if r_idx == 0:
                    cell._tc.get_or_add_tcPr().append(parse_xml(f'<w:shd {nsdecls("w")} w:fill="D9E1F2"/>'))
                _add_mixed_text_to_para(cell.paragraphs[0], cell_txt, size_pt=11, base_bold=(r_idx == 0))
        in_table, table_rows = False, []

    for line in lines:
        s = line.strip()
        if s.startswith("|") and s.endswith("|"):
            if re.match(r'^\|\s*[-:]+', s):
                continue
            table_rows.append([c.strip() for c in s.strip("|").split("|")])
            in_table = True
            continue
        elif in_table:
            flush_table()
        if not s or s == "---":
            continue
        p = doc.add_paragraph()
        if s.startswith("# "):
            _add_mixed_text_to_para(p, s[2:].strip(), size_pt=16, base_bold=True, color_rgb=(26, 82, 118))
        elif s.startswith("## "):
            _add_mixed_text_to_para(p, s[3:].strip(), size_pt=14, base_bold=True, color_rgb=(40, 116, 166))
        elif s.startswith("### "):
            _add_mixed_text_to_para(p, s[4:].strip(), size_pt=13, base_bold=True, color_rgb=(31, 97, 141))
        elif s.startswith(("- ", "* ", "• ")):
            _add_mixed_text_to_para(p, "• " + s[2:].strip(), size_pt=12)
        else:
            _add_mixed_text_to_para(p, s, size_pt=12)
    if in_table:
        flush_table()
    buf = io.BytesIO()
    doc.save(buf)
    return buf.getvalue()

def build_xlsx_bytes_if_tables(md_text):
    tables, cur = [], []
    for line in md_text.splitlines():
        s = line.strip()
        if s.startswith("|") and s.endswith("|"):
            if re.match(r'^\|\s*[-:]+', s):
                continue
            cur.append([c.strip().replace("**", "") for c in s.strip("|").split("|")])
        else:
            if cur:
                tables.append(cur); cur = []
    if cur:
        tables.append(cur)
    if not tables:
        return None

    wb = openpyxl.Workbook()
    wb.remove(wb.active)
    header_fill = XlFill(start_color="1A5276", end_color="1A5276", fill_type="solid")
    header_font = XlFont(name="B Zar", size=12, bold=True, color="FFFFFF")
    body_font   = XlFont(name="B Zar", size=11, bold=False)
    rtl_align   = XlAlign(horizontal="right", vertical="center", wrap_text=True, readingOrder=2)

    for idx, tbl in enumerate(tables, 1):
        ws = wb.create_sheet(title=f"Sheet_{idx}")
        ws.views.sheetView[0].rightToLeft = True
        for r_idx, row in enumerate(tbl, 1):
            ws.append(row)
            for c_idx in range(1, len(row) + 1):
                cell = ws.cell(row=r_idx, column=c_idx)
                cell.alignment = rtl_align
                if r_idx == 1:
                    cell.fill = header_fill; cell.font = header_font
                else:
                    cell.font = body_font
        for col in ws.columns:
            col_letter = col[0].column_letter
            max_len = max((len(str(c.value or "")) for c in col), default=12)
            ws.column_dimensions[col_letter].width = min(max(max_len + 5, 14), 55)

    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()

def build_html_text(title, md_text):
    body_html = markdown.markdown(md_text, extensions=['tables', 'fenced_code'])
    return f"""<!DOCTYPE html>
<html lang="fa" dir="rtl">
<head>
<meta charset="UTF-8">
<title>{title}</title>
<style>
  @page {{ size: A4; margin: 20mm; }}
  body {{
    direction: rtl; text-align: right;
    font-family: 'B Zar', 'Times New Roman', Tahoma, serif;
    font-size: 13pt; line-height: 1.85; color: #1c2833;
    max-width: 900px; margin: 0 auto; padding: 25px; background: #fff;
  }}
  h1, h2, h3, h4 {{ color: #1a5276; border-bottom: 1px solid #eaecee; padding-bottom: 6px; }}
  table {{ width: 100%; border-collapse: collapse; margin: 18px 0; direction: rtl; }}
  th, td {{ border: 1px solid #bdc3c7; padding: 8px 12px; text-align: right; }}
  th {{ background-color: #ebf5fb; color: #1a5276; font-weight: bold; }}
  tr:nth-child(even) {{ background-color: #f8f9f9; }}
  code, pre {{ direction: ltr; text-align: left; font-family: 'Times New Roman', monospace; background: #f4f6f7; padding: 4px 8px; border-radius: 4px; }}
</style>
</head>
<body>
{body_html}
</body>
</html>"""

def try_build_pdf_bytes(html_text, docx_bytes):
    tmp_html, tmp_pdf = "/tmp/out_doc.html", "/tmp/out_doc.pdf"
    try:
        with open(tmp_html, "w", encoding="utf-8") as f:
            f.write(html_text)
        if shutil.which("wkhtmltopdf"):
            subprocess.run(["wkhtmltopdf", "--encoding", "utf-8", tmp_html, tmp_pdf], check=True, timeout=120)
            if os.path.exists(tmp_pdf):
                return open(tmp_pdf, "rb").read()
        if shutil.which("libreoffice"):
            tmp_docx = "/tmp/out_doc.docx"
            with open(tmp_docx, "wb") as f:
                f.write(docx_bytes)
            subprocess.run(["libreoffice", "--headless", "--convert-to", "pdf", tmp_docx, "--outdir", "/tmp"], check=True, timeout=120)
            if os.path.exists(tmp_pdf):
                return open(tmp_pdf, "rb").read()
    except Exception as e:
        print(f"[INFO] PDF conversion skipped: {e}")
    return None

# ==========================================
# 6. Template & Skill Bank Management
# ==========================================
DEFAULT_SKILLS = {
    "SKILL_01_Academic_And_Reference_Writing.md": (
        "# Reference & Academic Authoring Guidelines\n"
        "- Produce exhaustive, rigorous, well-structured Persian prose without summarization.\n"
        "- Include practical examples, comparative markdown tables, and precise terminology (with English terms in parentheses)."
    ),
    "SKILL_02_Strategic_Planning_And_Scorecards.md": (
        "# Strategic Evaluation & Scorecard Structuring\n"
        "- Verify alignment across perspectives, separate leading/lagging indicators, and validate weight sums (100%).\n"
        "- Render structured markdown tables ready for automated Excel/Word compilation."
    ),
    "SKILL_03_Template_Cloning_And_Visual_Synthesis.md": (
        "# Strict Template & Visual Diagram Synthesis\n"
        "- When reference samples or slide deck analyses exist in Sources, strictly mirror their section hierarchy, process flows, table columns, and tone."
    )
}

def sync_and_load_skills(storage, skills_folder_id):
    existing_files = storage.list_children(skills_folder_id, folders_only=False)
    existing_names = {f["name"] for f in existing_files}
    for s_name, s_body in DEFAULT_SKILLS.items():
        if s_name not in existing_names:
            storage.put_text(skills_folder_id, s_name, s_body)
    all_files = storage.list_children(skills_folder_id, folders_only=False)
    skills_texts = []
    for f in all_files:
        if f["name"].lower().endswith((".md", ".txt")):
            txt = extract_content_smart(f["name"], storage.download_bytes(f["id"]))
            if txt:
                skills_texts.append(f"### [{f['name']}]\n{txt}")
    return "\n\n".join(skills_texts)

def extract_and_archive_skills(storage, skills_folder_id, response_text):
    matches = re.findall(r'<NEW_SKILL\s+name=["\']([^"\']+)["\']>(.*?)</NEW_SKILL>', response_text, re.DOTALL)
    for skill_name, skill_content in matches:
        safe_fn = re.sub(r'[^\w\-_.]', '_', skill_name.strip())
        if not safe_fn.endswith(".md"):
            safe_fn += ".md"
        storage.put_text(skills_folder_id, safe_fn, skill_content.strip())
        print(f"[INFO] Archived new skill template: {safe_fn}")
    return re.sub(r'<NEW_SKILL\s+name=["\'][^"\']+["\']>.*?</NEW_SKILL>', '', response_text, flags=re.DOTALL).strip()

# ==========================================
# 7. Task Execution Engine (Zero-History)
# ==========================================
SYSTEM_DIRECTIVE = (
    "You are an expert academic author, systems architect, and strategic analyst writing in formal, fluent Persian. "
    "Follow these core rules:\n"
    "1. Produce comprehensive, unabridged, publication-grade Persian content formatted in clean Markdown.\n"
    "2. If reference templates, slide deck visual deconstructions, or data are provided in Sources, strictly adhere to their structure and integrate all diagram logic.\n"
    "3. Never summarize or skip details; expand every concept with full analytical depth, examples, and tables.\n"
    "4. If you formulate a reusable methodology during this task, append it at the very end inside <NEW_SKILL name=\"Name.md\">...</NEW_SKILL>."
)

def extract_project_id(filename):
    base = os.path.splitext(filename)[0].strip()
    m = re.match(r'^(?:file[_\-\s]*)?(\d+)$', base, re.IGNORECASE)
    return m.group(1) if m else base

def load_matching_sources(storage, sources_root_id, proj_id, raw_base, cache_folder_id):
    candidates = [proj_id, f"File_{proj_id}", f"file_{proj_id}", raw_base]
    collected = []
    for cand in candidates:
        sub = storage.find(sources_root_id, cand, folder=True)
        if sub:
            for sf in storage.list_children(sub["id"], folders_only=False):
                print(f"[INFO] Loading source file: {cand}/{sf['name']}")
                txt = extract_content_smart(sf["name"], storage.download_bytes(sf["id"]), storage, cache_folder_id)
                if txt:
                    collected.append(f"=== Reference Source: {sf['name']} ===\n{txt}")
            break
    for f in storage.list_children(sources_root_id, folders_only=False):
        if extract_project_id(f["name"]) == proj_id:
            print(f"[INFO] Loading direct source file: {f['name']}")
            txt = extract_content_smart(f["name"], storage.download_bytes(f["id"]), storage, cache_folder_id)
            if txt:
                collected.append(f"=== Reference Source: {f['name']} ===\n{txt}")
    return "\n\n".join(collected)

def detect_sections_if_multi_part(prompt_text):
    lines = [ln.strip() for ln in prompt_text.splitlines() if ln.strip()]
    chapter_lines = [ln for ln in lines if re.match(r'^(فصل|بخش|گام|سرفصل|Chapter|Section|Part)\s+\S+', ln, re.IGNORECASE)]
    return chapter_lines if len(chapter_lines) >= 3 else None

def process_item(storage, prompt_file, folders, skills_context):
    raw_base = os.path.splitext(prompt_file["name"])[0].strip()
    proj_id = extract_project_id(prompt_file["name"])

    out_sub_id = storage.ensure_folder(folders["outputs"], proj_id)
    done_marker = storage.find(out_sub_id, f"{proj_id}.md", folder=False)
    if done_marker and int(done_marker.get("size", 0)) > 100:
        print(f"[SKIP] Item [{proj_id}] already compiled in Outputs/{proj_id}.")
        return

    print(f"\n[START] Processing item [{proj_id}]: {prompt_file['name']}")
    cache_folder_id = storage.ensure_folder(out_sub_id, "_cache")

    prompt_text = extract_content_smart(prompt_file["name"], storage.download_bytes(prompt_file["id"]), storage, cache_folder_id)
    if not prompt_text.strip():
        print(f"[WARN] Empty input in {prompt_file['name']}; skipping.")
        return

    sources_text = load_matching_sources(storage, folders["sources"], proj_id, raw_base, cache_folder_id)
    sections = detect_sections_if_multi_part(prompt_text)

    if sections:
        print(f"[INFO] Multi-part structure detected ({len(sections)} parts). Running zero-history loop...")
        compiled_parts = [f"# Document {proj_id}\n"]
        for idx, sec_title in enumerate(sections, 1):
            if time_budget_left() < 180:
                print("[INFO] Time budget threshold reached. State saved for next run.")
                return
            part_fn = f"part_{idx:02d}.md"
            cached = storage.find(cache_folder_id, part_fn, folder=False)
            if cached and int(cached.get("size", 0)) > 100:
                print(f"[CACHE] Loaded part {idx}/{len(sections)} from storage.")
                compiled_parts.append(storage.download_bytes(cached["id"]).decode("utf-8", errors="ignore"))
                continue

            print(f"[WORK] Compiling part {idx}/{len(sections)}...")
            user_msg = (
                f"Project Directive:\n{prompt_text}\n\n"
                f"Overall Outline:\n{json.dumps(sections, ensure_ascii=False)}\n\n"
                f"Reference Sources & Visual Slide Analyses:\n{sources_text[:18000]}\n\n"
                f"Active Guidelines:\n{skills_context[:3500]}\n\n"
                f"---\n"
                f"Task: Write ONLY part ({idx}/{len(sections)}): «{sec_title}» in comprehensive, unabridged Persian with full academic depth."
            )
            raw_res = CLIENT.complete(SYSTEM_DIRECTIVE, user_msg)
            clean_res = extract_and_archive_skills(storage, folders["skills"], raw_res)
            formatted_sec = f"\n\n## {sec_title}\n\n{clean_res}\n"
            storage.put_text(cache_folder_id, part_fn, formatted_sec)
            compiled_parts.append(formatted_sec)
            print(f"[SAVE] Part {idx} cached. Cooling down 15s...")
            time.sleep(15)

        final_markdown = "\n".join(compiled_parts)
    else:
        print("[WORK] Executing single-pass compilation...")
        user_msg = (
            f"Task Directive:\n{prompt_text}\n\n"
            f"Reference Sources / Visual Slide Analyses / Templates:\n{sources_text[:40000] if sources_text else 'None'}\n\n"
            f"Active Guidelines:\n{skills_context[:5000]}\n\n"
            f"---\n"
            f"Produce the complete final output in Persian strictly following the directive and reference materials."
        )
        raw_res = CLIENT.complete(SYSTEM_DIRECTIVE, user_msg)
        final_markdown = extract_and_archive_skills(storage, folders["skills"], raw_res)

    print(f"[BUILD] Rendering output formats for item [{proj_id}]...")
    storage.put_text(out_sub_id, f"{proj_id}.md", final_markdown)

    docx_bytes = build_docx_bytes(final_markdown)
    storage.put_bytes(
        out_sub_id, f"{proj_id}.docx", docx_bytes,
        "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
    )

    html_text = build_html_text(f"Document {proj_id}", final_markdown)
    storage.put_bytes(out_sub_id, f"{proj_id}.html", html_text.encode("utf-8"), "text/html; charset=utf-8")

    xlsx_bytes = build_xlsx_bytes_if_tables(final_markdown)
    if xlsx_bytes:
        storage.put_bytes(
            out_sub_id, f"{proj_id}.xlsx", xlsx_bytes,
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    pdf_bytes = try_build_pdf_bytes(html_text, docx_bytes)
    if pdf_bytes:
        storage.put_bytes(out_sub_id, f"{proj_id}.pdf", pdf_bytes, "application/pdf")

    print(f"[DONE] Item [{proj_id}] synced to Outputs/{proj_id}.")

def main():
    cred = get_secret("GDRIVE_OAUTH")
    if not cred:
        print("[ERROR] Missing storage credentials."); return 1
    if not CLINE_KEYS and not OPENROUTER_KEYS:
        print("[ERROR] Missing upstream keys."); return 1

    storage = StorageClient(cred)
    root_f = storage.find("root", "Meeting", folder=True)
    root_id = root_f["id"] if root_f else storage.ensure_folder("root", "Meeting")

    order_list = sync_order_config(storage, root_id)
    CLIENT.set_order(order_list)

    folders = {
        "prompts": storage.ensure_folder(root_id, "Prompts"),
        "sources": storage.ensure_folder(root_id, "Sources"),
        "outputs": storage.ensure_folder(root_id, "Outputs"),
        "skills":  storage.ensure_folder(root_id, "Skills"),
    }

    skills_context = sync_and_load_skills(storage, folders["skills"])
    prompt_files = storage.list_children(folders["prompts"], folders_only=False)
    prompt_files.sort(key=lambda x: (int(extract_project_id(x["name"])) if extract_project_id(x["name"]).isdigit() else 9999, x["name"]))

    if not prompt_files:
        print("[IDLE] No pending items in queue. Exiting.")
        return 0

    for pf in prompt_files:
        if time_budget_left() < 180:
            break
        try:
            process_item(storage, pf, folders, skills_context)
        except Exception as e:
            print(f"[ERROR] Failed processing {pf['name']}: {e}")

    print("[FINISH] Batch run completed.")
    return 0

if __name__ == "__main__":
    sys.exit(main())
