from flask import Flask, render_template, request, redirect, url_for, session, flash, send_file
from markupsafe import Markup, escape
from flask_sqlalchemy import SQLAlchemy
from sqlalchemy import inspect, text
from werkzeug.security import generate_password_hash, check_password_hash
from werkzeug.utils import secure_filename
from werkzeug.middleware.proxy_fix import ProxyFix
from docx import Document
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, Alignment, PatternFill
from authlib.integrations.flask_client import OAuth
from datetime import datetime
from zoneinfo import ZoneInfo
from io import BytesIO
import os, re, random, uuid, json, hashlib, subprocess, tempfile, shutil, platform, unicodedata
import fitz
from PIL import Image, ImageDraw

app = Flask(__name__)
app.wsgi_app = ProxyFix(app.wsgi_app, x_for=1, x_proto=1, x_host=1)

APP_TIMEZONE = os.environ.get('APP_TIMEZONE', 'Asia/Ho_Chi_Minh').strip() or 'Asia/Ho_Chi_Minh'
try:
    LOCAL_TZ = ZoneInfo(APP_TIMEZONE)
except Exception:
    LOCAL_TZ = ZoneInfo('Asia/Ho_Chi_Minh')

def local_now():
    # Giữ datetime dạng naive để tương thích dữ liệu các bản V6/V7,
    # nhưng giá trị luôn được tính theo múi giờ cấu hình của ứng dụng.
    return datetime.now(LOCAL_TZ).replace(tzinfo=None)

app.config['SECRET_KEY'] = os.environ.get('SECRET_KEY', 'lop-toan-online-v8-change-this')
app.config['SESSION_COOKIE_HTTPONLY'] = True
app.config['SESSION_COOKIE_SAMESITE'] = 'Lax'
app.config['SESSION_COOKIE_SECURE'] = os.environ.get('COOKIE_SECURE', '0') == '1'
app.config['PERMANENT_SESSION_LIFETIME'] = 60 * 60 * 12
app.config['MAX_CONTENT_LENGTH'] = int(os.environ.get('MAX_UPLOAD_MB', '50')) * 1024 * 1024
database_url = os.environ.get('DATABASE_URL', 'sqlite:///lop_toan_online_v6.db').strip()
if database_url.startswith('postgres://'):
    database_url = 'postgresql+psycopg://' + database_url[len('postgres://'):]
elif database_url.startswith('postgresql://'):
    database_url = 'postgresql+psycopg://' + database_url[len('postgresql://'):]
app.config['SQLALCHEMY_DATABASE_URI'] = database_url
app.config['SQLALCHEMY_TRACK_MODIFICATIONS'] = False
app.config['UPLOAD_FOLDER'] = 'uploads'
app.config['STATIC_UPLOAD_FOLDER'] = 'static/uploads'
os.makedirs(app.config['UPLOAD_FOLDER'], exist_ok=True)
os.makedirs(app.config['STATIC_UPLOAD_FOLDER'], exist_ok=True)
db = SQLAlchemy(app)

oauth = OAuth(app)
GOOGLE_CLIENT_ID = os.environ.get('GOOGLE_CLIENT_ID', '').strip()
GOOGLE_CLIENT_SECRET = os.environ.get('GOOGLE_CLIENT_SECRET', '').strip()
if GOOGLE_CLIENT_ID and GOOGLE_CLIENT_SECRET:
    oauth.register(
        name='google', client_id=GOOGLE_CLIENT_ID, client_secret=GOOGLE_CLIENT_SECRET,
        server_metadata_url='https://accounts.google.com/.well-known/openid-configuration',
        client_kwargs={'scope': 'openid email profile'}
    )

class Classroom(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    name = db.Column(db.String(80), unique=True, nullable=False)
    grade = db.Column(db.String(20), default='')
    description = db.Column(db.Text, default='')

class User(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    username = db.Column(db.String(80), unique=True, nullable=False)
    email = db.Column(db.String(160), unique=True, nullable=True)
    password_hash = db.Column(db.String(255), nullable=False)
    full_name = db.Column(db.String(150), nullable=False)
    role = db.Column(db.String(20), nullable=False)
    classroom_id = db.Column(db.Integer, db.ForeignKey('classroom.id'))
    avatar = db.Column(db.String(255), default='')

class SiteSetting(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    owner_id = db.Column(db.Integer, db.ForeignKey('user.id'), unique=True, nullable=False)
    site_title = db.Column(db.String(150), default='Lớp Toán Online')
    teacher_label = db.Column(db.String(150), default='Giáo viên Toán')
    welcome_text = db.Column(db.Text, default='Học Toán mỗi ngày – tiến bộ mỗi ngày!')

class Lesson(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    title = db.Column(db.String(200), nullable=False)
    description = db.Column(db.Text, default='')
    content = db.Column(db.Text, default='')
    resource_url = db.Column(db.String(500), default='')
    classroom_id = db.Column(db.Integer, db.ForeignKey('classroom.id'), nullable=False)
    created_by = db.Column(db.Integer, db.ForeignKey('user.id'), nullable=False)
    is_published = db.Column(db.Boolean, default=False)
    created_at = db.Column(db.DateTime, default=datetime.utcnow)

class BankQuestion(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    owner_id = db.Column(db.Integer, db.ForeignKey('user.id'), nullable=False)
    grade = db.Column(db.String(20), default='')
    topic = db.Column(db.String(120), default='')
    difficulty = db.Column(db.String(20), default='Trung bình')
    domain = db.Column(db.String(20), default='Đại số')
    qtype = db.Column(db.String(20), nullable=False)
    content = db.Column(db.Text, nullable=False)
    option_a = db.Column(db.Text, default='')
    option_b = db.Column(db.Text, default='')
    option_c = db.Column(db.Text, default='')
    option_d = db.Column(db.Text, default='')
    correct_answer = db.Column(db.String(10), default='')
    explanation = db.Column(db.Text, default='')
    points = db.Column(db.Float, default=1.0)  # trọng số
    image_path = db.Column(db.String(255), default='')

class Assignment(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    title = db.Column(db.String(200), nullable=False)
    description = db.Column(db.Text, default='')
    created_by = db.Column(db.Integer, db.ForeignKey('user.id'), nullable=False)
    duration_minutes = db.Column(db.Integer, default=45)
    starts_at = db.Column(db.DateTime, nullable=True)
    due_at = db.Column(db.DateTime, nullable=True)
    score_scale = db.Column(db.Float, default=10.0)
    shuffle_questions = db.Column(db.Boolean, default=False)
    shuffle_options = db.Column(db.Boolean, default=False)
    show_result = db.Column(db.Boolean, default=True)
    allow_retake = db.Column(db.Boolean, default=False)
    is_published = db.Column(db.Boolean, default=False)
    created_at = db.Column(db.DateTime, default=datetime.utcnow)

class AssignmentClassroom(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    assignment_id = db.Column(db.Integer, db.ForeignKey('assignment.id'), nullable=False)
    classroom_id = db.Column(db.Integer, db.ForeignKey('classroom.id'), nullable=False)
    __table_args__ = (db.UniqueConstraint('assignment_id', 'classroom_id', name='uq_assignment_classroom'),)

class AssignmentQuestion(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    assignment_id = db.Column(db.Integer, db.ForeignKey('assignment.id'), nullable=False)
    question_id = db.Column(db.Integer, db.ForeignKey('bank_question.id'), nullable=False)
    order_no = db.Column(db.Integer, default=0)

class Submission(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    assignment_id = db.Column(db.Integer, db.ForeignKey('assignment.id'), nullable=False)
    student_id = db.Column(db.Integer, db.ForeignKey('user.id'), nullable=False)
    started_at = db.Column(db.DateTime, default=datetime.utcnow)
    submitted_at = db.Column(db.DateTime)
    submitted = db.Column(db.Boolean, default=False)
    auto_score = db.Column(db.Float, default=0.0)
    manual_score = db.Column(db.Float, default=0.0)
    max_score = db.Column(db.Float, default=10.0)
    @property
    def total_score(self):
        return round((self.auto_score or 0) + (self.manual_score or 0), 2)

class Answer(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    submission_id = db.Column(db.Integer, db.ForeignKey('submission.id'), nullable=False)
    question_id = db.Column(db.Integer, db.ForeignKey('bank_question.id'), nullable=False)
    answer_text = db.Column(db.Text, default='')
    auto_score = db.Column(db.Float, default=0.0)
    manual_score = db.Column(db.Float, default=0.0)
    teacher_note = db.Column(db.Text, default='')

def me():
    uid = session.get('user_id')
    return db.session.get(User, uid) if uid else None

def teacher_only():
    u = me(); return bool(u and u.role == 'teacher')

def student_only():
    u = me(); return bool(u and u.role == 'student')

def parse_dt(v):
    if not v: return None
    try: return datetime.strptime(v, '%Y-%m-%dT%H:%M')
    except: return None

def my_setting():
    u = me()
    if not u or u.role != 'teacher': return None
    s = SiteSetting.query.filter_by(owner_id=u.id).first()
    if not s:
        s = SiteSetting(owner_id=u.id, teacher_label=u.full_name)
        db.session.add(s); db.session.commit()
    return s

def assignment_class_ids(a):
    return [x.classroom_id for x in AssignmentClassroom.query.filter_by(assignment_id=a.id).all()]

def assignment_classes(a):
    ids = assignment_class_ids(a)
    return Classroom.query.filter(Classroom.id.in_(ids)).order_by(Classroom.name).all() if ids else []

def assignment_class_names(a):
    return ', '.join(c.name for c in assignment_classes(a)) or 'Chưa chọn lớp'

def assignment_status(a):
    if not a.is_published:
        return ('Bản nháp', 'draft')
    now = local_now()
    if a.starts_at and now < a.starts_at:
        return ('Chưa mở', 'scheduled')
    if a.due_at and now > a.due_at:
        return ('Đã kết thúc', 'ended')
    return ('Đang diễn ra', 'active')

def student_has_assignment(a, classroom_id):
    if not classroom_id: return False
    return AssignmentClassroom.query.filter_by(assignment_id=a.id, classroom_id=classroom_id).first() is not None

def set_assignment_classes(a, class_ids):
    AssignmentClassroom.query.filter_by(assignment_id=a.id).delete()
    seen = set()
    for sid in class_ids:
        try: cid = int(sid)
        except: continue
        if cid in seen or not db.session.get(Classroom, cid): continue
        seen.add(cid)
        db.session.add(AssignmentClassroom(assignment_id=a.id, classroom_id=cid))

def question_score_map(a, questions):
    valid = [q for q in questions if q]
    weights = [max(0.0, float(q.points or 0)) for q in valid]
    total = sum(weights)
    if total <= 0:
        weights = [1.0 for _ in valid]; total = float(len(valid) or 1)
    scale = max(0.01, float(a.score_scale or 10.0))
    return {q.id: round(scale * w / total, 4) for q, w in zip(valid, weights)}

def save_uploaded_image(f, prefix='q'):
    if not f or not getattr(f, 'filename', ''): return ''
    fn = secure_filename(f.filename)
    ext = os.path.splitext(fn)[1].lower()
    if ext not in ['.png', '.jpg', '.jpeg', '.webp', '.gif']:
        return ''
    name = f'{prefix}_{uuid.uuid4().hex[:12]}{ext}'
    f.save(os.path.join(app.config['STATIC_UPLOAD_FOLDER'], name))
    return name

def xml_text(element):
    out = []
    for node in element.iter():
        if node.tag.split('}')[-1] == 't' and node.text:
            out.append(node.text)
    return ''.join(out).strip()

def save_docx_image(part):
    ctype = getattr(part, 'content_type', '') or ''
    ext_map = {'image/png': '.png', 'image/jpeg': '.jpg', 'image/gif': '.gif', 'image/webp': '.webp', 'image/x-emf': '.emf', 'image/x-wmf': '.wmf'}
    ext = ext_map.get(ctype, os.path.splitext(str(getattr(part, 'partname', '')))[1] or '.png')
    # Trình duyệt không đọc EMF/WMF tốt; file mẫu MathType có preview PNG/JPG nên ưu tiên phần đó.
    name = f'word_{uuid.uuid4().hex[:12]}{ext}'
    with open(os.path.join(app.config['STATIC_UPLOAD_FOLDER'], name), 'wb') as out:
        out.write(part.blob)
    return name

def paragraph_rich_text(doc, para, image_cache):
    """Giữ đúng thứ tự chữ + ảnh preview của công thức MathType/Equation trong Word."""
    parts = []
    for node in para._p.iter():
        local = node.tag.split('}')[-1]
        if local == 't' and node.text:
            parts.append(node.text)
        elif local in ('imagedata', 'blip'):
            rid = None
            for k, v in node.attrib.items():
                if k.endswith('}id') or k.endswith('}embed'):
                    rid = v; break
            if not rid or rid not in doc.part.related_parts:
                continue
            relpart = doc.part.related_parts[rid]
            ctype = getattr(relpart, 'content_type', '') or ''
            if not ctype.startswith('image/'):
                continue
            if rid not in image_cache:
                image_cache[rid] = save_docx_image(relpart)
            parts.append(f'[[img:{image_cache[rid]}]]')
    return ''.join(parts).strip()

def clean_level_tag(text):
    # Giữ nhãn NB/TH/VD/VDC trong nội dung để giáo viên nhìn được mức độ.
    return text.strip()

def parse_tf_answer_text(text):
    """Trả về danh sách Đ/S theo thứ tự xuất hiện trong một dòng đáp án."""
    x = text.lower().replace('đúng', 'đ').replace('sai', 's')
    vals = []
    # ưu tiên các cặp a) Đ, b) S... nếu có
    for m in re.finditer(r'(?:[abcd]\s*[).:]\s*)?([đs])(?:\b|\s|,|$)', x, re.I):
        v = m.group(1).lower()
        vals.append('A' if v == 'đ' else 'B')
    return vals

def _find_office_converter():
    """Tìm LibreOffice/soffice trên Linux hoặc Windows."""
    for name in ('libreoffice', 'soffice'):
        exe = shutil.which(name)
        if exe:
            return exe
    if os.name == 'nt':
        for exe in (
            r'C:\Program Files\LibreOffice\program\soffice.exe',
            r'C:\Program Files (x86)\LibreOffice\program\soffice.exe',
        ):
            if os.path.exists(exe):
                return exe
    return None


def _docx_to_pdf(path, out_dir):
    """Chuyển DOCX thành PDF bằng Word trên Windows hoặc LibreOffice.

    Mục đích của bước này là render công thức MathType/Equation đúng bố cục Word,
    thay vì tự ghép từng ảnh công thức nhỏ (dễ bị chồng ký tự).
    """
    os.makedirs(out_dir, exist_ok=True)
    pdf_path = os.path.join(out_dir, os.path.splitext(os.path.basename(path))[0] + '.pdf')

    # 1) Nếu chạy trên Windows và máy có Microsoft Word: dùng Word COM để giữ bố cục tốt nhất.
    if os.name == 'nt':
        try:
            import win32com.client  # pywin32 chỉ cài trên Windows
            word = win32com.client.DispatchEx('Word.Application')
            word.Visible = False
            doc = word.Documents.Open(os.path.abspath(path), ReadOnly=True)
            doc.SaveAs(os.path.abspath(pdf_path), FileFormat=17)  # wdFormatPDF
            doc.Close(False)
            word.Quit()
            if os.path.exists(pdf_path):
                return pdf_path
        except Exception:
            try:
                word.Quit()
            except Exception:
                pass

    # 2) Linux/Render hoặc Windows có LibreOffice.
    office = _find_office_converter()
    if office:
        try:
            subprocess.run(
                [office, '--headless', '--convert-to', 'pdf', '--outdir', out_dir, os.path.abspath(path)],
                stdout=subprocess.PIPE, stderr=subprocess.PIPE, timeout=120, check=False
            )
            if os.path.exists(pdf_path):
                return pdf_path
        except Exception:
            pass
    return None


def _pdf_lines(pdf):
    rows = []
    for pno, page in enumerate(pdf):
        data = page.get_text('dict')
        for block in data.get('blocks', []):
            if block.get('type') != 0:
                continue
            for line in block.get('lines', []):
                spans = line.get('spans', [])
                text = ''.join(sp.get('text', '') for sp in spans).strip()
                if not text:
                    continue
                bbox = line.get('bbox')
                if not bbox and spans:
                    xs0=[sp['bbox'][0] for sp in spans]; ys0=[sp['bbox'][1] for sp in spans]
                    xs1=[sp['bbox'][2] for sp in spans]; ys1=[sp['bbox'][3] for sp in spans]
                    bbox=(min(xs0),min(ys0),max(xs1),max(ys1))
                if bbox:
                    rows.append({'page':pno, 'bbox':tuple(bbox), 'text':text})
    return rows


def _student_render_copy(path, out_path):
    """Tạo bản DOCX chỉ để render cho học sinh.

    Gỡ toàn bộ underline vì nhiều ngân hàng Word dùng gạch chân để đánh dấu đáp án đúng.
    Dữ liệu đáp án vẫn được đọc từ file gốc trước khi render.
    """
    try:
        d = Document(path)
        for p in d.paragraphs:
            for r in p.runs:
                if r.underline:
                    r.underline = False
        for table in d.tables:
            for row in table.rows:
                for cell in row.cells:
                    for p in cell.paragraphs:
                        for r in p.runs:
                            if r.underline:
                                r.underline = False
        d.save(out_path)
        return out_path
    except Exception:
        return path


def _question_answer_metadata(doc):
    """Đọc đáp án theo nhiều kiểu Word thường gặp.

    Hỗ trợ:
    - đáp án gạch chân trong lựa chọn A/B/C/D;
    - bảng đáp án dạng hàng 'Câu' / 'Đáp án';
    - phần 'BẢNG ĐÁP ÁN...' với các dòng 'Câu 16: 100';
    - câu đúng/sai có các ý a,b,c,d gạch chân để đánh dấu ý đúng.
    """
    answer_map = {}
    under_mcq = {}
    under_tf = {}
    in_key = False
    current_q = None

    for p in doc.paragraphs:
        txt = p.text.strip()
        if re.match(r'^BẢNG\s+ĐÁP\s+ÁN', txt, re.I):
            in_key = True
            current_q = None
            continue
        qm = re.match(r'^Câu\s*(\d+)\s*[:.)-]', txt, re.I)
        if qm and not in_key:
            current_q = int(qm.group(1))
        elif qm and in_key:
            # Dòng đáp án ngắn kiểu: Câu 16: 100
            mm = re.match(r'^Câu\s*(\d+)\s*[:.)-]\s*(.+?)\s*$', txt, re.I)
            if mm:
                val = mm.group(2).strip()
                if val and not re.match(r'^\[?(NB|TH|VD|VDC)\]?', val, re.I):
                    answer_map[int(mm.group(1))] = val
            continue

        if in_key:
            continue
        if current_q is None:
            continue

        # Tìm ký tự/nhãn được gạch chân trong file giáo viên.
        under_text = ' '.join((r.text or '') for r in p.runs if r.underline and (r.text or '').strip())
        if under_text:
            # Lựa chọn A-D gạch chân. Chỉ lấy nhãn lựa chọn, không lấy chữ cái nằm trong từ.
            for m in re.finditer(r'(?<![A-Za-zÀ-ỹ])([ABCD])\s*[.):;]?', under_text, re.I):
                under_mcq.setdefault(current_q, m.group(1).upper())
            # Mệnh đề a)-d) gạch chân = mệnh đề đúng.
            for m in re.finditer(r'(?<![A-Za-zÀ-ỹ])([abcd])\s*\)', under_text):
                under_tf.setdefault(current_q, set()).add(m.group(1).lower())

    # Bảng đáp án: hàng đầu là Câu, hàng sau là Đáp án.
    for table in doc.tables:
        rows = [[c.text.strip() for c in row.cells] for row in table.rows]
        for i in range(len(rows)-1):
            head, ans = rows[i], rows[i+1]
            if not head or not ans:
                continue
            if re.match(r'^Câu$', head[0], re.I) and re.match(r'^Đáp\s*án$', ans[0], re.I):
                for qtxt, atxt in zip(head[1:], ans[1:]):
                    qm = re.search(r'\d+', qtxt)
                    if qm and atxt:
                        answer_map[int(qm.group())] = atxt.strip()

    # Underline là nguồn dự phòng, bảng/đáp án tường minh được ưu tiên.
    for q, a in under_mcq.items():
        answer_map.setdefault(q, a)
    return answer_map, under_tf


def _answer_to_mcq_letter(v):
    x = (v or '').strip().upper().replace('ĐÚNG','Đ').replace('SAI','S')
    if x in ('A','B','C','D'):
        return x
    if x in ('Đ','D','TRUE','T'):
        return 'A'
    if x in ('S','FALSE','F'):
        return 'B'
    return ''


def render_word_question_images(path):
    """Render từng câu Word thành PNG, che đáp án và giữ đúng công thức/hình.

    Bản V6.5 còn gỡ underline trước khi render vì nhiều file ngân hàng dùng underline
    để đánh dấu đáp án đúng. Nhờ vậy học sinh không nhìn thấy dấu đáp án.
    """
    digest = hashlib.sha256(b'v66-boundary-crop-3|' + open(path,'rb').read()).hexdigest()[:10]
    cached = {}
    prefix = f'wordq_{digest}_q'
    try:
        for fn in os.listdir(app.config['STATIC_UPLOAD_FOLDER']):
            if fn.startswith(prefix) and fn.endswith('.png'):
                m = re.match(rf'^wordq_{digest}_q(\d+)\.png$', fn)
                if m: cached[int(m.group(1))] = fn
    except Exception:
        pass
    if cached:
        return cached

    with tempfile.TemporaryDirectory(prefix='loptoan_word_') as td:
        safe_docx = _student_render_copy(path, os.path.join(td, 'student_render.docx'))
        pdf_path = _docx_to_pdf(safe_docx, td)
        if not pdf_path:
            return {}
        try:
            pdf = fitz.open(pdf_path)
        except Exception:
            return {}
        lines = _pdf_lines(pdf)

        # Không render phần bảng đáp án ở cuối file.
        key_row = next((r for r in lines if re.match(r'^\s*BẢNG\s+ĐÁP\s+ÁN', r['text'], re.I)), None)
        key_pos = (key_row['page'], key_row['bbox'][1]) if key_row else None

        starts = []
        for row in lines:
            if key_pos and (row['page'] > key_pos[0] or (row['page'] == key_pos[0] and row['bbox'][1] >= key_pos[1])):
                continue
            m = re.match(r'^\s*Câu\s*(\d+)\s*[\.:)]', row['text'], re.I)
            if m:
                starts.append({'qnum':int(m.group(1)), 'page':row['page'], 'y':row['bbox'][1]})
        uniq, seen = [], set()
        for st in sorted(starts, key=lambda x:(x['page'],x['y'])):
            if st['qnum'] not in seen:
                uniq.append(st); seen.add(st['qnum'])
        starts = uniq
        if not starts:
            pdf.close(); return {}

        # Dữ kiện chung cho nhiều câu: giữ cả bảng/hình trước câu đầu và ghép vào từng câu trong khoảng.
        shared = []
        for row in lines:
            mm = re.search(r'Dữ\s*kiện\s*chung\s*cho\s*các\s*câu\s*từ\s*(\d+)\s*đến\s*(\d+)', row['text'], re.I)
            if mm:
                q1, q2 = int(mm.group(1)), int(mm.group(2))
                first = next((x for x in starts if x['qnum'] == q1), None)
                if first:
                    shared.append({'q1':q1,'q2':q2,'page':row['page'],'y':row['bbox'][1],
                                   'end_page':first['page'],'end_y':first['y']})

        answer_rows = [r for r in lines if re.match(r'^\s*(Đáp\s*án|ĐA)\s*[:\-]', r['text'], re.I)]
        result = {}
        zoom = 2.0

        def render_region(sp, sy, ep, ey):
            pieces=[]
            for pno in range(sp, ep+1):
                page=pdf[pno]
                top=max(0, sy-12) if pno==sp else 0
                bottom=min(page.rect.height, ey-12) if pno==ep else page.rect.height
                if bottom <= top+2: continue
                clip=fitz.Rect(0,top,page.rect.width,bottom)
                pix=page.get_pixmap(matrix=fitz.Matrix(zoom,zoom),clip=clip,alpha=False)
                im=Image.frombytes('RGB',[pix.width,pix.height],pix.samples)
                try:
                    gp=im.convert('L'); mp=gp.point(lambda px: 255 if px < 248 else 0); bb=mp.getbbox()
                    if bb:
                        pp=12; im=im.crop((max(0,bb[0]-pp),max(0,bb[1]-pp),min(im.width,bb[2]+pp),min(im.height,bb[3]+pp)))
                except Exception:
                    pass
                pieces.append(im)
            if not pieces: return None
            width=max(x.width for x in pieces); height=sum(x.height for x in pieces)
            merged=Image.new('RGB',(width,height),'white'); y=0
            for im in pieces: merged.paste(im,(0,y)); y+=im.height
            return merged

        for i, st in enumerate(starts):
            nxt = starts[i+1] if i+1 < len(starts) else None
            end_page = nxt['page'] if nxt else len(pdf)-1
            end_y = nxt['y'] if nxt else pdf[end_page].rect.height
            if key_pos:
                kp, ky = key_pos
                if kp < end_page or (kp == end_page and ky < end_y):
                    end_page, end_y = kp, ky

            pieces=[]
            sh = next((x for x in shared if x['q1'] <= st['qnum'] <= x['q2']), None)
            if sh:
                sim=render_region(sh['page'], sh['y'], sh['end_page'], sh['end_y'])
                if sim: pieces.append(sim)

            for pno in range(st['page'], end_page+1):
                page = pdf[pno]
                top = max(0, st['y'] - 2) if pno == st['page'] else 0
                bottom = max(top+2, end_y - 2) if pno == end_page else page.rect.height
                if bottom <= top + 2: continue
                clip = fitz.Rect(0, top, page.rect.width, bottom)
                pix = page.get_pixmap(matrix=fitz.Matrix(zoom, zoom), clip=clip, alpha=False)
                im = Image.frombytes('RGB', [pix.width, pix.height], pix.samples)
                draw = ImageDraw.Draw(im)
                for ar in answer_rows:
                    if ar['page'] != pno: continue
                    x0,y0,x1,y1 = ar['bbox']
                    if y1 < top or y0 > bottom: continue
                    draw.rectangle((max(0,int((x0-3)*zoom)), max(0,int((y0-top-3)*zoom)),
                                    im.width, min(im.height,int((y1-top+5)*zoom))), fill='white')

                # Với câu kéo dài qua ngắt trang, mỗi mảnh trang có thể chứa một khoảng
                # trắng rất lớn. Cắt trắng TỪNG MẢNH trước khi ghép để câu 19/20 kiểu
                # ngân hàng Toán 8 không còn dính câu trước hoặc có khoảng trắng khổng lồ.
                try:
                    gray_piece = im.convert('L')
                    mask_piece = gray_piece.point(lambda px: 255 if px < 248 else 0)
                    bb = mask_piece.getbbox()
                    if bb:
                        pad_piece = 12
                        px0=max(0,bb[0]-pad_piece); py0=max(0,bb[1]-pad_piece)
                        px1=min(im.width,bb[2]+pad_piece); py1=min(im.height,bb[3]+pad_piece)
                        im=im.crop((px0,py0,px1,py1))
                except Exception:
                    pass
                pieces.append(im)

            if not pieces: continue
            width=max(x.width for x in pieces); height=sum(x.height for x in pieces)
            merged=Image.new('RGB',(width,height),'white'); y=0
            for im in pieces: merged.paste(im,(0,y)); y+=im.height

            # Cắt khoảng trắng thừa nhưng giữ một lề nhỏ. Việc này chỉ dựa trên pixel
            # nên không làm mất công thức/hình vẽ như cách phân tích text/OCR.
            try:
                gray = merged.convert('L')
                # mọi pixel tối hơn 248 được coi là nội dung; ảnh nền nhạt/hình học vẫn giữ lại.
                mask = gray.point(lambda px: 255 if px < 248 else 0)
                bbox = mask.getbbox()
                if bbox:
                    pad = 18
                    x0=max(0,bbox[0]-pad); y0=max(0,bbox[1]-pad)
                    x1=min(merged.width,bbox[2]+pad); y1=min(merged.height,bbox[3]+pad)
                    merged=merged.crop((x0,y0,x1,y1))
            except Exception:
                pass

            if merged.width > 1800:
                ratio=1800/merged.width
                merged=merged.resize((1800,max(1,int(merged.height*ratio))),Image.Resampling.LANCZOS)
            fn=f'wordq_{digest}_q{st["qnum"]}.png'
            merged.save(os.path.join(app.config['STATIC_UPLOAD_FOLDER'],fn),'PNG',optimize=True)
            result[st['qnum']]=fn
        pdf.close()
        return result


def parse_word(path):
    """Bộ nhập Word tổng quát V6.5.

    Nhận các mẫu đã gặp trong 3 ngân hàng của người dùng:
    A/B/C/D, Đúng/Sai từng câu, Đúng/Sai nhiều mệnh đề, điền khuyết,
    trả lời ngắn, đáp án gạch chân, bảng đáp án và đáp án tường minh.
    """
    doc = Document(path)
    answer_map, under_tf = _question_answer_metadata(doc)
    image_cache = {}
    blocks, cur = [], None
    section = ''
    in_answer_key = False

    for para in doc.paragraphs:
        text = paragraph_rich_text(doc, para, image_cache)
        plain = re.sub(r'\[\[img:[^\]]+\]\]', '', text).strip()
        if re.match(r'^BẢNG\s+ĐÁP\s+ÁN', plain, re.I):
            if cur: blocks.append(cur); cur=None
            in_answer_key = True
            continue
        if in_answer_key:
            continue
        if re.match(r'^PHẦN\s+[IVX]+', plain, re.I):
            section = plain
            continue
        if re.match(r'^Câu\s*\d+\s*[:.)\-]', plain, re.I):
            if cur: blocks.append(cur)
            cur = {'lines':[text], 'section':section}
        elif cur and text:
            cur['lines'].append(text)
    if cur: blocks.append(cur)

    rendered = render_word_question_images(path)
    out=[]
    for b in blocks:
        lines=b['lines']; sec=b.get('section','')
        first_plain=re.sub(r'\[\[img:[^\]]+\]\]','',lines[0])
        mnum=re.match(r'^Câu\s*(\d+)\s*[:.)\-]\s*', first_plain, re.I)
        qnum=int(mnum.group(1)) if mnum else 0
        content=re.sub(r'^Câu\s*\d+\s*[:.)\-]\s*','',lines[0],flags=re.I)
        qimage=rendered.get(qnum,'')

        # Thu đáp án tường minh trong block trước, ưu tiên hơn bảng/underline.
        explicit_ans=''
        weight=1.0; explanation=''
        for line in lines[1:]:
            plain=re.sub(r'\[\[img:[^\]]+\]\]','',line).strip()
            am=re.match(r'^(Đáp\s*án|ĐA)\s*[:\-]\s*(.*)$', plain, re.I)
            if am and am.group(2).strip(): explicit_ans=am.group(2).strip()
            wm=re.match(r'^(Điểm|Trọng\s*số)\s*[:\-]\s*([\d.,]+)', plain, re.I)
            if wm:
                try: weight=float(wm.group(2).replace(',','.'))
                except: pass
            em=re.match(r'^(Lời\s*giải|Giải\s*thích)\s*[:\-]\s*(.*)$', plain, re.I)
            if em: explanation=em.group(2).strip()
        raw_ans=explicit_ans or answer_map.get(qnum,'')

        # Dạng Đúng/Sai nhiều mệnh đề a)-d).
        statement_rows=[]
        for line in lines[1:]:
            plain=re.sub(r'\[\[img:[^\]]+\]\]','',line).strip()
            sm=re.match(r'^([abcd])\)\s*(.*)$', plain, re.I)
            if sm: statement_rows.append({'label':sm.group(1).lower(),'text':sm.group(2).strip()})
        sec_tf = bool(re.search(r'ĐÚNG\s*/?\s*SAI|ĐÚNG\s*SAI', sec, re.I))

        if len(statement_rows) >= 2:
            vals=parse_tf_answer_text(raw_ans)
            true_labels=under_tf.get(qnum,set())
            for idx, strow in enumerate(statement_rows):
                if idx < len(vals): corr=vals[idx]
                elif true_labels: corr='A' if strow['label'] in true_labels else 'B'
                else: corr=''
                if not corr: continue
                display_content=(f'Câu {qnum} — Mệnh đề {strow["label"]})' if qimage
                                 else f'{content}\n{strow["label"]}) {strow["text"]}')
                out.append(dict(qtype='mcq',content=display_content,
                                opts={'A':'Đúng','B':'Sai','C':'','D':''},correct=corr,
                                points=weight,explanation=explanation or f'Câu {qnum}, mệnh đề {strow["label"]}',
                                image_path=qimage,source_qnum=qnum,source_kind='tf'))
            if statement_rows and (vals or true_labels):
                continue

        # Dạng mỗi Câu là một mệnh đề Đúng/Sai (như ngân hàng Toán 8).
        tf_letter=_answer_to_mcq_letter(raw_ans)
        if sec_tf and tf_letter:
            out.append(dict(qtype='mcq',content=(f'Câu {qnum} — Chọn Đúng hoặc Sai theo đề trong hình.' if qimage else clean_level_tag(content)),
                            opts={'A':'Đúng','B':'Sai','C':'','D':''},correct=tf_letter,points=weight,
                            explanation=explanation,image_path=qimage,source_qnum=qnum,source_kind='tf_single'))
            continue

        # Nhận diện câu trả lời ngắn theo tên phần hoặc đáp án không phải A-D.
        short_section=bool(re.search(r'ĐIỀN\s*KHUYẾT|TRẢ\s*LỜI\s*NGẮN', sec, re.I))
        mcq_letter=_answer_to_mcq_letter(raw_ans)

        # Nhận diện có lựa chọn A-D kể cả nhiều lựa chọn nằm chung một paragraph.
        all_plain='\n'.join(re.sub(r'\[\[img:[^\]]+\]\]','',x) for x in lines[1:])
        choice_letters=set(m.group(1).upper() for m in re.finditer(r'(?<![A-Za-zÀ-ỹ])([ABCD])\s*[.):-]', all_plain, re.I))
        has_choices=len(choice_letters)>=2

        opts={k:'' for k in 'ABCD'}
        for line in lines[1:]:
            plain=re.sub(r'\[\[img:[^\]]+\]\]','',line).strip()
            # Tách được các lựa chọn đơn dòng; nếu chung dòng thì ảnh render sẽ hiển thị đầy đủ.
            om=re.match(r'^([ABCD])[.):-]\s*(.*)$', plain, re.I)
            if om: opts[om.group(1).upper()]=om.group(2).strip()

        if short_section and raw_ans:
            qtype='short'; correct=str(raw_ans).strip(); skind='short'
        elif (has_choices or mcq_letter) and mcq_letter:
            qtype='mcq'; correct=mcq_letter; skind='mcq'
            # Với ảnh render, chỉ cần bốn nút chữ cái; nội dung lựa chọn nằm đúng trong ảnh Word.
            if qimage:
                opts={'A':'A','B':'B','C':'C','D':'D'}
        elif raw_ans and not mcq_letter:
            qtype='short'; correct=str(raw_ans).strip(); skind='short'
        else:
            qtype='essay'; correct=''; skind='essay'

        if qimage:
            if qtype=='mcq': display_content=f'Câu {qnum} — Chọn đáp án đúng trong hình.'
            elif qtype=='short': display_content=f'Câu {qnum} — Nhập đáp án theo yêu cầu trong hình.'
            else: display_content=f'Câu {qnum} — Trình bày bài làm theo đề trong hình.'
        else:
            display_content=clean_level_tag(content)
        out.append(dict(qtype=qtype,content=display_content,opts=opts,correct=correct,
                        points=weight,explanation=explanation,image_path=qimage,
                        source_qnum=qnum,source_kind=skind))
    return out

def normalize_short_answer(v):
    v = (v or '').strip().lower().replace('−','-').replace('–','-')
    v = re.sub(r'\s+', '', v)
    return v

def render_rich(v):
    """Hiển thị token ảnh nội tuyến [[img:file]] an toàn trong nội dung câu hỏi/đáp án."""
    if not v:
        return Markup('')
    text = str(v)
    chunks = re.split(r'(\[\[img:[^\]]+\]\])', text)
    out = []
    for ch in chunks:
        m = re.fullmatch(r'\[\[img:([^\]]+)\]\]', ch)
        if m:
            fn = os.path.basename(m.group(1))
            out.append(f'<img class="eqimg" src="/static/uploads/{fn}" alt="công thức">')
        else:
            out.append(str(escape(ch)).replace('\n','<br>'))
    return Markup(''.join(out))


K12_SAMPLE_SHA256 = "5eb53256639a4777c240d238cbb921cd7cec2399e25135b19665ec314c7c4e43"

def load_k12_fraction_sample():
    path = os.path.join(os.path.dirname(__file__), 'k12_fraction_sample.json')
    with open(path, 'r', encoding='utf-8') as f:
        return json.load(f)

def import_k12_items_to_assignment(a, grade='6', topic='Phân số'):
    data = load_k12_fraction_sample(); order = AssignmentQuestion.query.filter_by(assignment_id=a.id).count(); count = 0
    for x in data.get('items', []):
        q = BankQuestion(owner_id=me().id, grade=grade or data.get('grade','6'), topic=topic or data.get('topic','Phân số'),
                         qtype=x['qtype'], content=x['content'], option_a=x.get('option_a',''), option_b=x.get('option_b',''),
                         option_c=x.get('option_c',''), option_d=x.get('option_d',''), correct_answer=x.get('correct_answer',''),
                         explanation=x.get('explanation',''), points=float(x.get('points',1) or 1), image_path=x.get('image_path',''))
        db.session.add(q); db.session.flush(); order += 1; count += 1
        db.session.add(AssignmentQuestion(assignment_id=a.id, question_id=q.id, order_no=order))
    db.session.commit(); return count

app.jinja_env.filters['rich'] = render_rich

def repair_k12_fraction_equations():
    """Sửa các câu 13/14 đã nạp từ V6.2 để không phải tạo lại đề hoặc mất dữ liệu."""
    fixes = [
        ('k12ps_056.png', '[TH] Kết quả của phép tính (xem công thức dưới đây) là', 'q13_stem.png'),
        ('k12ps_060.png', '[TH] Kết quả của phép tính (xem công thức dưới đây) là', 'q14_stem.png'),
    ]
    changed = False
    for token, content, img in fixes:
        for q in BankQuestion.query.filter(BankQuestion.content.contains(token)).all():
            q.content = content
            q.image_path = img
            changed = True
    if changed:
        db.session.commit()


@app.context_processor
def ctx():
    u = me(); setting = None
    if u and u.role == 'teacher': setting = SiteSetting.query.filter_by(owner_id=u.id).first()
    return {'me': u, 'site_setting': setting, 'google_enabled': bool(GOOGLE_CLIENT_ID and GOOGLE_CLIENT_SECRET),
            'now': datetime.now(), 'assignment_class_names': assignment_class_names, 'assignment_status': assignment_status}

@app.route('/')
def index():
    u = me()
    if not u: return redirect(url_for('login'))
    return redirect(url_for('teacher_dashboard' if u.role == 'teacher' else 'student_dashboard'))

@app.route('/login', methods=['GET', 'POST'])
def login():
    if request.method == 'POST':
        u = User.query.filter_by(username=request.form['username'].strip()).first()
        if not u or not check_password_hash(u.password_hash, request.form['password']):
            flash('Sai tài khoản hoặc mật khẩu.', 'error'); return render_template('login.html')
        session['user_id'] = u.id; return redirect(url_for('index'))
    return render_template('login.html')

@app.route('/login/google')
def login_google():
    if not GOOGLE_CLIENT_ID or not GOOGLE_CLIENT_SECRET:
        flash('Đăng nhập Google chưa được cấu hình. Xem file HUONG_DAN_GOOGLE.txt.', 'error'); return redirect(url_for('login'))
    return oauth.google.authorize_redirect(url_for('google_callback', _external=True))

@app.route('/auth/google/callback')
def google_callback():
    if not GOOGLE_CLIENT_ID or not GOOGLE_CLIENT_SECRET: return redirect(url_for('login'))
    token = oauth.google.authorize_access_token(); info = token.get('userinfo') or oauth.google.userinfo()
    email = (info.get('email') or '').lower().strip()
    u = User.query.filter(db.func.lower(User.email) == email).first() if email else None
    if not u:
        flash('Email Google này chưa được giáo viên gắn với tài khoản trong hệ thống.', 'error'); return redirect(url_for('login'))
    session['user_id'] = u.id; return redirect(url_for('index'))

@app.route('/logout')
def logout():
    session.clear(); return redirect(url_for('login'))

@app.route('/teacher')
def teacher_dashboard():
    if not teacher_only(): return redirect(url_for('login'))
    u = me(); classes = Classroom.query.order_by(Classroom.name).all()
    assignments = Assignment.query.filter_by(created_by=u.id).order_by(Assignment.id.desc()).all()
    students = User.query.filter_by(role='student').all()
    published = sum(a.is_published for a in assignments)
    submitted = Submission.query.join(Assignment, Submission.assignment_id == Assignment.id).filter(Assignment.created_by == u.id, Submission.submitted == True).count()
    lessons = Lesson.query.filter_by(created_by=u.id).count(); questions = BankQuestion.query.filter_by(owner_id=u.id).count()
    return render_template('teacher_dashboard.html', classes=classes, assignments=assignments, students=students,
                           published=published, submitted=submitted, lessons=lessons, questions=questions)

@app.route('/teacher/settings', methods=['GET', 'POST'])
def settings():
    if not teacher_only(): return redirect(url_for('login'))
    s = my_setting(); u = me()
    if request.method == 'POST':
        s.site_title = request.form.get('site_title', 'Lớp Toán Online').strip() or 'Lớp Toán Online'
        s.teacher_label = request.form.get('teacher_label', u.full_name).strip() or u.full_name
        s.welcome_text = request.form.get('welcome_text', '').strip(); u.full_name = s.teacher_label
        u.email = request.form.get('email', '').strip().lower() or None
        f = request.files.get('avatar')
        if f and f.filename:
            name = save_uploaded_image(f, f'avatar_{u.id}')
            if name: u.avatar = name
        db.session.commit(); flash('Đã lưu thông tin hệ thống.', 'ok'); return redirect(url_for('settings'))
    return render_template('settings.html', setting=s)

@app.route('/teacher/classes', methods=['GET', 'POST'])
def classes():
    if not teacher_only(): return redirect(url_for('login'))
    if request.method == 'POST':
        name = request.form['name'].strip()
        if Classroom.query.filter_by(name=name).first(): flash('Lớp này đã tồn tại.', 'error')
        else:
            db.session.add(Classroom(name=name, grade=request.form.get('grade', '').strip(), description=request.form.get('description', '').strip()))
            db.session.commit(); flash('Đã tạo lớp.', 'ok')
        return redirect(url_for('classes'))
    rows = []
    for c in Classroom.query.order_by(Classroom.name).all():
        acount = AssignmentClassroom.query.filter_by(classroom_id=c.id).count()
        rows.append((c, User.query.filter_by(role='student', classroom_id=c.id).count(), acount))
    return render_template('classes.html', rows=rows)

def _slug_username(text):
    text = unicodedata.normalize('NFD', str(text or ''))
    text = ''.join(ch for ch in text if unicodedata.category(ch) != 'Mn')
    text = text.lower().replace('đ', 'd')
    text = re.sub(r'[^a-z0-9]+', '', text)
    return text or 'hocsinh'

def _unique_username(full_name, class_name=''):
    parts = [x for x in re.split(r'\s+', str(full_name or '').strip()) if x]
    base_name = (parts[-1] if parts else 'hocsinh') + ''.join(x[0] for x in parts[:-1])
    base = _slug_username(base_name + str(class_name or ''))[:60] or 'hocsinh'
    candidate = base; n = 2
    while User.query.filter_by(username=candidate).first():
        candidate = (base[:55] + str(n))
        n += 1
    return candidate

@app.route('/teacher/students', methods=['GET', 'POST'])
def students():
    if not teacher_only(): return redirect(url_for('login'))
    if request.method == 'POST':
        username = request.form['username'].strip(); email = request.form.get('email', '').strip().lower() or None
        if User.query.filter_by(username=username).first(): flash('Tên đăng nhập đã tồn tại.', 'error')
        elif email and User.query.filter(db.func.lower(User.email) == email).first(): flash('Email đã được dùng.', 'error')
        else:
            db.session.add(User(username=username, email=email, password_hash=generate_password_hash(request.form['password']),
                                full_name=request.form['full_name'].strip(), role='student', classroom_id=int(request.form['classroom_id'])))
            db.session.commit(); flash('Đã tạo tài khoản học sinh.', 'ok')
        return redirect(url_for('students'))
    return render_template('students.html', students=User.query.filter_by(role='student').order_by(User.full_name).all(), classes=Classroom.query.order_by(Classroom.name).all())

@app.route('/teacher/students/excel-template')
def student_excel_template():
    if not teacher_only(): return redirect(url_for('login'))
    path = os.path.join(app.root_path, 'mau_de', 'Mau_Danh_Sach_Hoc_Sinh.xlsx')
    return send_file(path, as_attachment=True, download_name='Mau_Danh_Sach_Hoc_Sinh.xlsx',
                     mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')

@app.route('/teacher/students/import-excel', methods=['POST'])
def import_students_excel():
    if not teacher_only(): return redirect(url_for('login'))
    f = request.files.get('excel_file')
    if not f or not getattr(f, 'filename', '') or not f.filename.lower().endswith('.xlsx'):
        flash('Vui lòng chọn file Excel .xlsx.', 'error'); return redirect(url_for('students'))
    auto_create_class = bool(request.form.get('auto_create_class'))
    try:
        wb = load_workbook(f, data_only=True)
        ws = wb['Danh sách học sinh'] if 'Danh sách học sinh' in wb.sheetnames else wb.active
    except Exception as e:
        flash('Không đọc được file Excel: ' + str(e), 'error'); return redirect(url_for('students'))

    header_map = {}
    aliases = {
        'full_name': {'họ và tên','ho va ten','họ tên','ho ten','tên học sinh','ten hoc sinh'},
        'class_name': {'lớp','lop','lớp học','lop hoc'},
        'username': {'tên đăng nhập','ten dang nhap','tài khoản','tai khoan','username'},
        'password': {'mật khẩu','mat khau','password'},
        'email': {'email google','email','gmail'}
    }
    def norm(v):
        t = unicodedata.normalize('NFD', str(v or '').strip().lower())
        t = ''.join(ch for ch in t if unicodedata.category(ch) != 'Mn').replace('đ','d')
        return re.sub(r'\s+', ' ', t)
    for cell in ws[1]:
        hv = norm(cell.value)
        for key, vals in aliases.items():
            if hv in {norm(x) for x in vals}: header_map[key] = cell.column
    if 'full_name' not in header_map or 'class_name' not in header_map:
        flash('File Excel phải có ít nhất 2 cột: Họ và tên, Lớp.', 'error'); return redirect(url_for('students'))

    results=[]; created=0; skipped=0
    for row in range(2, ws.max_row + 1):
        full_name = str(ws.cell(row, header_map['full_name']).value or '').strip()
        class_name = str(ws.cell(row, header_map['class_name']).value or '').strip()
        username = str(ws.cell(row, header_map.get('username', 0)).value or '').strip() if header_map.get('username') else ''
        password = str(ws.cell(row, header_map.get('password', 0)).value or '').strip() if header_map.get('password') else ''
        email = str(ws.cell(row, header_map.get('email', 0)).value or '').strip().lower() if header_map.get('email') else ''
        if not any([full_name, class_name, username, email]): continue
        if not full_name or not class_name:
            results.append([row, full_name, class_name, username, password, email, 'Bỏ qua: thiếu Họ và tên hoặc Lớp']); skipped += 1; continue
        classroom = Classroom.query.filter(db.func.lower(Classroom.name) == class_name.lower()).first()
        if not classroom and auto_create_class:
            m = re.search(r'([6-9])', class_name)
            classroom = Classroom(name=class_name, grade=m.group(1) if m else '', description='Tạo tự động khi nhập Excel')
            db.session.add(classroom); db.session.flush()
        if not classroom:
            results.append([row, full_name, class_name, username, password, email, 'Bỏ qua: lớp chưa tồn tại']); skipped += 1; continue
        if not username: username = _unique_username(full_name, class_name)
        if not password: password = '123456'
        if User.query.filter_by(username=username).first():
            results.append([row, full_name, class_name, username, password, email, 'Bỏ qua: tên đăng nhập đã tồn tại']); skipped += 1; continue
        if email and User.query.filter(db.func.lower(User.email) == email).first():
            results.append([row, full_name, class_name, username, password, email, 'Bỏ qua: email đã tồn tại']); skipped += 1; continue
        db.session.add(User(username=username, email=email or None, password_hash=generate_password_hash(password),
                            full_name=full_name, role='student', classroom_id=classroom.id))
        results.append([row, full_name, class_name, username, password, email, 'Đã tạo'])
        created += 1
    db.session.commit()

    out = Workbook(); rws = out.active; rws.title = 'Kết quả nhập học sinh'
    rws.append(['Dòng Excel','Họ và tên','Lớp','Tên đăng nhập','Mật khẩu','Email Google','Trạng thái'])
    for x in results: rws.append(x)
    for cell in rws[1]:
        cell.font = Font(bold=True, color='FFFFFF'); cell.fill = PatternFill('solid', fgColor='153D8A'); cell.alignment = Alignment(horizontal='center')
    widths = [12,28,14,22,16,30,36]
    for i,w in enumerate(widths,1): rws.column_dimensions[chr(64+i)].width = w
    for row in rws.iter_rows():
        for cell in row: cell.alignment = Alignment(vertical='top', wrap_text=True)
    bio = BytesIO(); out.save(bio); bio.seek(0)
    flash(f'Đã nhập {created} học sinh; bỏ qua {skipped} dòng. File kết quả đang được tải xuống.', 'ok')
    return send_file(bio, as_attachment=True, download_name=f'Ket_qua_nhap_hoc_sinh_{datetime.now().strftime("%Y%m%d_%H%M%S")}.xlsx',
                     mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')

@app.route('/teacher/student/<int:user_id>/reset', methods=['POST'])
def reset_password(user_id):
    if not teacher_only(): return redirect(url_for('login'))
    u = db.session.get(User, user_id)
    if u and u.role == 'student':
        u.password_hash = generate_password_hash(request.form.get('password', '123456')); db.session.commit(); flash('Đã đổi mật khẩu học sinh.', 'ok')
    return redirect(url_for('students'))

@app.route('/teacher/student/<int:user_id>/delete', methods=['POST'])
def delete_student(user_id):
    if not teacher_only(): return redirect(url_for('login'))
    u = db.session.get(User, user_id)
    if u and u.role == 'student':
        for s in Submission.query.filter_by(student_id=u.id).all():
            Answer.query.filter_by(submission_id=s.id).delete(); db.session.delete(s)
        db.session.delete(u); db.session.commit(); flash('Đã xóa học sinh.', 'ok')
    return redirect(url_for('students'))

@app.route('/teacher/lessons', methods=['GET', 'POST'])
def lessons():
    if not teacher_only(): return redirect(url_for('login'))
    if request.method == 'POST':
        db.session.add(Lesson(title=request.form['title'].strip(), description=request.form.get('description', '').strip(),
                              content=request.form.get('content', '').strip(), resource_url=request.form.get('resource_url', '').strip(),
                              classroom_id=int(request.form['classroom_id']), created_by=me().id, is_published=bool(request.form.get('is_published'))))
        db.session.commit(); flash('Đã tạo bài giảng.', 'ok'); return redirect(url_for('lessons'))
    return render_template('lessons.html', lessons=Lesson.query.filter_by(created_by=me().id).order_by(Lesson.id.desc()).all(), classes=Classroom.query.order_by(Classroom.name).all())

@app.route('/teacher/lesson/<int:lid>/delete', methods=['POST'])
def delete_lesson(lid):
    if not teacher_only(): return redirect(url_for('login'))
    x = db.session.get(Lesson, lid)
    if x and x.created_by == me().id:
        db.session.delete(x); db.session.commit(); flash('Đã xóa bài giảng.', 'ok')
    return redirect(url_for('lessons'))

@app.route('/teacher/word-template')
def download_word_template():
    if not teacher_only(): return redirect(url_for('login'))
    path = os.path.join(os.path.dirname(__file__), 'mau_de', 'Mau_De_Toan_TracNghiem_TuLuan.docx')
    if not os.path.exists(path):
        flash('Không tìm thấy file Word mẫu đi kèm phần mềm.', 'error')
        return redirect(request.referrer or url_for('teacher_dashboard'))
    return send_file(path, as_attachment=True, download_name='Mau_De_Toan_TracNghiem_TuLuan.docx',
                     mimetype='application/vnd.openxmlformats-officedocument.wordprocessingml.document')

@app.route('/teacher/question-bank', methods=['GET', 'POST'])
def question_bank():
    if not teacher_only(): return redirect(url_for('login'))
    if request.method == 'POST':
        image_path = save_uploaded_image(request.files.get('image_file'), 'question')
        try: weight = max(0, float(request.form.get('points', 1) or 1))
        except: weight = 1.0
        q = BankQuestion(owner_id=me().id, grade=request.form.get('grade', '').strip(), topic=request.form.get('topic', '').strip(),
                         difficulty=request.form.get('difficulty', 'Trung bình').strip() or 'Trung bình', domain=request.form.get('domain', 'Đại số').strip() or 'Đại số', qtype=request.form['qtype'], content=request.form['content'].strip(), option_a=request.form.get('option_a', '').strip(),
                         option_b=request.form.get('option_b', '').strip(), option_c=request.form.get('option_c', '').strip(),
                         option_d=request.form.get('option_d', '').strip(), correct_answer=request.form.get('correct_answer', '').strip().upper(),
                         explanation=request.form.get('explanation', '').strip(), points=weight, image_path=image_path)
        db.session.add(q); db.session.commit(); flash('Đã thêm câu hỏi.', 'ok'); return redirect(url_for('question_bank'))
    grade = request.args.get('grade', '').strip()
    qtype = request.args.get('qtype', '').strip()
    difficulty = request.args.get('difficulty', '').strip()
    domain = request.args.get('domain', '').strip()
    topic = request.args.get('topic', '').strip()
    search = request.args.get('search', '').strip()
    try: page = max(1, int(request.args.get('page', '1')))
    except: page = 1
    per_page = 100
    query = BankQuestion.query.filter_by(owner_id=me().id)
    if grade: query = query.filter(BankQuestion.grade == grade)
    if qtype: query = query.filter(BankQuestion.qtype == qtype)
    if difficulty: query = query.filter(BankQuestion.difficulty == difficulty)
    if domain: query = query.filter(BankQuestion.domain == domain)
    if topic: query = query.filter(BankQuestion.topic.ilike(f'%{topic}%'))
    if search: query = query.filter(BankQuestion.content.ilike(f'%{search}%'))
    total = query.count()
    qs = query.order_by(BankQuestion.id.desc()).offset((page-1)*per_page).limit(per_page).all()
    topic_rows = db.session.query(BankQuestion.topic).filter_by(owner_id=me().id).distinct().order_by(BankQuestion.topic).all()
    topics = [x[0] for x in topic_rows if x[0]]
    return render_template('question_bank.html', questions=qs, total=total, page=page, per_page=per_page,
                           filter_grade=grade, filter_qtype=qtype, filter_difficulty=difficulty, filter_domain=domain, filter_topic=topic, filter_search=search, topics=topics)

@app.route('/teacher/question-bank/load-1000', methods=['POST'])
def load_builtin_1000_questions():
    if not teacher_only(): return redirect(url_for('login'))
    data_path = os.path.join(os.path.dirname(__file__), 'question_bank_1000.json')
    if not os.path.exists(data_path):
        flash('Không tìm thấy gói 1.000 câu hỏi đi kèm phần mềm.', 'error')
        return redirect(url_for('question_bank'))
    with open(data_path, 'r', encoding='utf-8') as f:
        items = json.load(f)
    existing = {x[0] for x in db.session.query(BankQuestion.explanation).filter(
        BankQuestion.owner_id == me().id, BankQuestion.explanation.like('§NH1000§%')).all()}
    added = 0
    for idx, x in enumerate(items, start=1):
        marker = f"§NH1000§{x.get('grade','')}-{idx:04d}"
        explanation = marker + (' | ' + x.get('explanation','') if x.get('explanation') else '')
        if explanation in existing: continue
        db.session.add(BankQuestion(owner_id=me().id, grade=x.get('grade',''), topic=x.get('topic',''),
            difficulty=x.get('difficulty','Trung bình'), domain=x.get('domain','Đại số'), qtype=x.get('qtype','mcq'), content=x.get('content',''), option_a=x.get('option_a',''),
            option_b=x.get('option_b',''), option_c=x.get('option_c',''), option_d=x.get('option_d',''),
            correct_answer=x.get('correct_answer',''), explanation=explanation, points=float(x.get('points',1) or 1), image_path=''))
        added += 1
    db.session.commit()
    flash(f'Đã nạp thêm {added} câu. Ngân hàng tích hợp gồm 1.000 câu Toán lớp 6–9; bấm lại sẽ không nhân đôi.', 'ok')
    return redirect(url_for('question_bank'))

@app.route('/teacher/question-bank/upload-word', methods=['POST'])
def bank_upload_word():
    if not teacher_only(): return redirect(url_for('login'))
    f = request.files.get('word_file')
    if not f or not f.filename.lower().endswith('.docx'):
        flash('Hãy chọn file Word .docx', 'error'); return redirect(url_for('question_bank'))
    fn = f'{uuid.uuid4().hex[:8]}_{secure_filename(f.filename)}'; path = os.path.join(app.config['UPLOAD_FOLDER'], fn); f.save(path)
    items = parse_word(path)
    for x in items:
        db.session.add(BankQuestion(owner_id=me().id, grade=request.form.get('grade', ''), topic=request.form.get('topic', ''),
            difficulty=request.form.get('difficulty', 'Trung bình'), domain=request.form.get('domain', 'Đại số'), qtype=x['qtype'], content=x['content'], option_a=x['opts']['A'], option_b=x['opts']['B'], option_c=x['opts']['C'], option_d=x['opts']['D'],
            correct_answer=x['correct'], explanation=x['explanation'], points=x['points'], image_path=x['image_path']))
    db.session.commit(); rendered_count=sum(1 for x in items if x.get('image_path')); flash(f'Đã nhập {len(items)} mục từ Word. {rendered_count} mục được render theo bố cục Word; trắc nghiệm có đáp án sẽ tự chấm.', 'ok'); return redirect(url_for('question_bank'))

@app.route('/teacher/question/<int:qid>/delete', methods=['POST'])
def delete_bank_question(qid):
    if not teacher_only(): return redirect(url_for('login'))
    q = db.session.get(BankQuestion, qid)
    if q and q.owner_id == me().id:
        AssignmentQuestion.query.filter_by(question_id=q.id).delete(); db.session.delete(q); db.session.commit(); flash('Đã xóa câu hỏi.', 'ok')
    return redirect(url_for('question_bank'))

@app.route('/teacher/assignments/new', methods=['GET', 'POST'])
def new_assignment():
    if not teacher_only(): return redirect(url_for('login'))
    classes_all = Classroom.query.order_by(Classroom.name).all()
    if request.method == 'POST':
        class_ids = request.form.getlist('classroom_ids')
        if not class_ids:
            flash('Hãy chọn ít nhất một lớp để giao bài.', 'error'); return render_template('new_assignment.html', classes=classes_all)
        try: scale = max(0.01, float(request.form.get('score_scale', 10) or 10))
        except: scale = 10.0
        a = Assignment(title=request.form['title'].strip(), description=request.form.get('description', '').strip(), created_by=me().id,
                       duration_minutes=max(1, int(request.form.get('duration_minutes', 45))), starts_at=parse_dt(request.form.get('starts_at')),
                       due_at=parse_dt(request.form.get('due_at')), score_scale=scale, shuffle_questions=bool(request.form.get('shuffle_questions')),
                       shuffle_options=bool(request.form.get('shuffle_options')), show_result=bool(request.form.get('show_result')),
                       allow_retake=bool(request.form.get('allow_retake')), is_published=bool(request.form.get('is_published')))
        if a.starts_at and a.due_at and a.starts_at >= a.due_at:
            flash('Thời gian kết thúc phải sau thời gian bắt đầu.', 'error'); return render_template('new_assignment.html', classes=classes_all)
        db.session.add(a); db.session.flush(); set_assignment_classes(a, class_ids); db.session.commit()
        flash('Đã tạo bài kiểm tra. Bạn có thể chọn câu hỏi hoặc nhập trực tiếp từ Word.', 'ok'); return redirect(url_for('edit_assignment', assignment_id=a.id))
    return render_template('new_assignment.html', classes=classes_all)

@app.route('/teacher/assignment/<int:assignment_id>', methods=['GET', 'POST'])
def edit_assignment(assignment_id):
    if not teacher_only(): return redirect(url_for('login'))
    a = db.session.get(Assignment, assignment_id)
    if not a or a.created_by != me().id: return 'Không có quyền', 403
    if request.method == 'POST':
        ids = request.form.getlist('question_ids'); existing = {x.question_id for x in AssignmentQuestion.query.filter_by(assignment_id=a.id).all()}
        order = AssignmentQuestion.query.filter_by(assignment_id=a.id).count()
        for sid in ids:
            qid = int(sid)
            if qid not in existing:
                order += 1; db.session.add(AssignmentQuestion(assignment_id=a.id, question_id=qid, order_no=order))
        db.session.commit(); flash('Đã thêm câu hỏi vào đề.', 'ok')
    aq = AssignmentQuestion.query.filter_by(assignment_id=a.id).order_by(AssignmentQuestion.order_no).all()
    selected = [db.session.get(BankQuestion, x.question_id) for x in aq]
    bank_grade = request.args.get('bank_grade', '').strip()
    bank_qtype = request.args.get('bank_qtype', '').strip()
    bank_difficulty = request.args.get('bank_difficulty', '').strip()
    bank_domain = request.args.get('bank_domain', '').strip()
    bank_topic = request.args.get('bank_topic', '').strip()
    bank_search = request.args.get('bank_search', '').strip()
    bq = BankQuestion.query.filter_by(owner_id=me().id)
    if bank_grade: bq = bq.filter(BankQuestion.grade == bank_grade)
    if bank_qtype: bq = bq.filter(BankQuestion.qtype == bank_qtype)
    if bank_difficulty: bq = bq.filter(BankQuestion.difficulty == bank_difficulty)
    if bank_domain: bq = bq.filter(BankQuestion.domain == bank_domain)
    if bank_topic: bq = bq.filter(BankQuestion.topic.ilike(f'%{bank_topic}%'))
    if bank_search: bq = bq.filter(BankQuestion.content.ilike(f'%{bank_search}%'))
    bank = bq.order_by(BankQuestion.id.desc()).limit(120).all()
    return render_template('edit_assignment.html', assignment=a, selected=selected, bank=bank,
                           classes=Classroom.query.order_by(Classroom.name).all(), selected_class_ids=assignment_class_ids(a),
                           score_map=question_score_map(a, selected), bank_grade=bank_grade, bank_qtype=bank_qtype,
                           bank_difficulty=bank_difficulty, bank_domain=bank_domain, bank_topic=bank_topic, bank_search=bank_search)

@app.route('/teacher/assignment/<int:assignment_id>/auto-generate', methods=['POST'])
def auto_generate_assignment(assignment_id):
    if not teacher_only(): return redirect(url_for('login'))
    a = db.session.get(Assignment, assignment_id)
    if not a or a.created_by != me().id: return 'Không có quyền', 403
    grade = request.form.get('auto_grade', '').strip()
    preset = request.form.get('auto_preset', 'Trung bình').strip()
    domain = request.form.get('auto_domain', '').strip()
    topic = request.form.get('auto_topic', '').strip()
    try: count = max(1, min(100, int(request.form.get('auto_count', 20) or 20)))
    except: count = 20
    if not grade:
        flash('Hãy chọn khối lớp để tạo đề tự động.', 'error')
        return redirect(url_for('edit_assignment', assignment_id=a.id))

    # Tỉ lệ mức độ cho ba kiểu đề. Phần dư được bù từ các mức còn nguồn câu.
    ratios = {
        'Dễ': [('Dễ', .70), ('Trung bình', .30), ('Khó', 0)],
        'Trung bình': [('Dễ', .30), ('Trung bình', .50), ('Khó', .20)],
        'Khó': [('Dễ', .10), ('Trung bình', .30), ('Khó', .60)],
    }
    plan = ratios.get(preset, ratios['Trung bình'])
    existing = {x.question_id for x in AssignmentQuestion.query.filter_by(assignment_id=a.id).all()}
    base = BankQuestion.query.filter_by(owner_id=me().id, grade=grade)
    if domain: base = base.filter(BankQuestion.domain == domain)
    if topic: base = base.filter(BankQuestion.topic.ilike(f'%{topic}%'))
    available = [q for q in base.all() if q.id not in existing]
    if not available:
        flash('Không tìm thấy câu phù hợp. Hãy nạp ngân hàng 1.000 câu hoặc đổi bộ lọc.', 'error')
        return redirect(url_for('edit_assignment', assignment_id=a.id))

    wanted = {}
    used = 0
    for i,(level,ratio) in enumerate(plan):
        n = count-used if i == len(plan)-1 else int(round(count*ratio))
        n = max(0,n); wanted[level]=n; used += n
    chosen=[]
    rng=random.SystemRandom()
    for level,_ in plan:
        pool=[q for q in available if q.difficulty==level and q not in chosen]
        n=min(wanted[level],len(pool))
        if n: chosen.extend(rng.sample(pool,n))
    if len(chosen)<count:
        pool=[q for q in available if q not in chosen]
        n=min(count-len(chosen),len(pool))
        if n: chosen.extend(rng.sample(pool,n))
    rng.shuffle(chosen)
    order=AssignmentQuestion.query.filter_by(assignment_id=a.id).count()
    for q in chosen:
        order += 1
        db.session.add(AssignmentQuestion(assignment_id=a.id, question_id=q.id, order_no=order))
    db.session.commit()
    counts={d:sum(1 for q in chosen if q.difficulty==d) for d in ['Dễ','Trung bình','Khó']}
    flash(f"Đã tạo đề {preset.lower()} gồm {len(chosen)} câu: {counts['Dễ']} dễ, {counts['Trung bình']} trung bình, {counts['Khó']} khó.", 'ok')
    return redirect(url_for('edit_assignment', assignment_id=a.id))

@app.route('/teacher/assignment/<int:assignment_id>/upload-word', methods=['POST'])
def assignment_upload_word(assignment_id):
    if not teacher_only(): return redirect(url_for('login'))
    a = db.session.get(Assignment, assignment_id)
    if not a or a.created_by != me().id: return 'Không có quyền', 403
    f = request.files.get('word_file')
    if not f or not f.filename.lower().endswith('.docx'):
        flash('Hãy chọn file Word .docx.', 'error'); return redirect(url_for('edit_assignment', assignment_id=a.id))
    fn = f'{uuid.uuid4().hex[:8]}_{secure_filename(f.filename)}'; path = os.path.join(app.config['UPLOAD_FOLDER'], fn); f.save(path)
    items = parse_word(path); order = AssignmentQuestion.query.filter_by(assignment_id=a.id).count()
    for x in items:
        q = BankQuestion(owner_id=me().id, grade=request.form.get('grade', '').strip(), topic=request.form.get('topic', '').strip(),
                         difficulty=request.form.get('difficulty', 'Trung bình').strip() or 'Trung bình', domain=request.form.get('domain', 'Đại số').strip() or 'Đại số', qtype=x['qtype'], content=x['content'], option_a=x['opts']['A'], option_b=x['opts']['B'], option_c=x['opts']['C'], option_d=x['opts']['D'],
                         correct_answer=x['correct'], explanation=x['explanation'], points=x['points'], image_path=x['image_path'])
        db.session.add(q); db.session.flush(); order += 1
        db.session.add(AssignmentQuestion(assignment_id=a.id, question_id=q.id, order_no=order))
    db.session.commit(); rendered_count=sum(1 for x in items if x.get('image_path')); flash(f'Đã tạo đề từ Word: thêm {len(items)} mục; {rendered_count} mục giữ nguyên công thức bằng ảnh render. Các câu có đáp án được tự chấm.', 'ok')
    return redirect(url_for('edit_assignment', assignment_id=a.id))

@app.route('/teacher/assignment/<int:assignment_id>/import-k12-fraction-sample', methods=['POST'])
def import_k12_fraction_sample_route(assignment_id):
    if not teacher_only(): return redirect(url_for('login'))
    a = db.session.get(Assignment, assignment_id)
    if not a or a.created_by != me().id: return 'Không có quyền', 403
    sample_path = os.path.join(os.path.dirname(__file__), 'samples', 'CauHoi_PhanSo_CoDapAn_K12Online.docx')
    if not os.path.exists(sample_path):
        flash('Không tìm thấy file đề mẫu đi kèm.', 'error'); return redirect(url_for('edit_assignment', assignment_id=a.id))
    items = parse_word(sample_path); order = AssignmentQuestion.query.filter_by(assignment_id=a.id).count()
    for x in items:
        q = BankQuestion(owner_id=me().id, grade='6', topic='Phân số', difficulty='Trung bình', domain='Đại số', qtype=x['qtype'], content=x['content'],
                         option_a=x['opts']['A'], option_b=x['opts']['B'], option_c=x['opts']['C'], option_d=x['opts']['D'],
                         correct_answer=x['correct'], explanation=x['explanation'], points=x['points'], image_path=x['image_path'])
        db.session.add(q); db.session.flush(); order += 1
        db.session.add(AssignmentQuestion(assignment_id=a.id, question_id=q.id, order_no=order))
    db.session.commit()
    rendered_count = sum(1 for x in items if x.get('image_path'))
    flash(f'Đã tích hợp đề Phân số mẫu: {len(items)} mục tự chấm; {rendered_count} mục dùng ảnh render Word để giữ nguyên công thức.', 'ok')
    return redirect(url_for('edit_assignment', assignment_id=a.id))

@app.route('/teacher/assignment/<int:assignment_id>/settings', methods=['POST'])
def assignment_settings(assignment_id):
    if not teacher_only(): return redirect(url_for('login'))
    a = db.session.get(Assignment, assignment_id)
    if not a or a.created_by != me().id: return 'Không có quyền', 403
    class_ids = request.form.getlist('classroom_ids')
    if not class_ids:
        flash('Hãy chọn ít nhất một lớp.', 'error'); return redirect(url_for('edit_assignment', assignment_id=a.id))
    starts_at = parse_dt(request.form.get('starts_at')); due_at = parse_dt(request.form.get('due_at'))
    if starts_at and due_at and starts_at >= due_at:
        flash('Thời gian kết thúc phải sau thời gian bắt đầu.', 'error'); return redirect(url_for('edit_assignment', assignment_id=a.id))
    try: scale = max(0.01, float(request.form.get('score_scale', 10) or 10))
    except: scale = 10.0
    a.title = request.form['title'].strip(); a.description = request.form.get('description', '').strip()
    a.duration_minutes = max(1, int(request.form.get('duration_minutes', 45))); a.starts_at = starts_at; a.due_at = due_at; a.score_scale = scale
    a.shuffle_questions = bool(request.form.get('shuffle_questions')); a.shuffle_options = bool(request.form.get('shuffle_options'))
    a.show_result = bool(request.form.get('show_result')); a.allow_retake = bool(request.form.get('allow_retake'))
    set_assignment_classes(a, class_ids); db.session.commit(); flash('Đã lưu cấu hình.', 'ok'); return redirect(url_for('edit_assignment', assignment_id=a.id))

@app.route('/teacher/assignment/<int:assignment_id>/publish', methods=['POST'])
def publish_assignment(assignment_id):
    if not teacher_only(): return redirect(url_for('login'))
    a = db.session.get(Assignment, assignment_id)
    if not a or a.created_by != me().id: return 'Không có quyền', 403
    if not assignment_class_ids(a):
        flash('Chưa thể giao bài: hãy chọn ít nhất một lớp.', 'error')
        return redirect(request.referrer or url_for('edit_assignment', assignment_id=a.id))
    qcount = AssignmentQuestion.query.filter_by(assignment_id=a.id).count()
    if qcount <= 0:
        flash('Chưa thể giao bài: đề chưa có câu hỏi.', 'error')
        return redirect(request.referrer or url_for('edit_assignment', assignment_id=a.id))
    now = local_now()
    if a.due_at and a.due_at <= now:
        flash('Chưa thể giao bài: thời gian kết thúc đã qua. Hãy sửa lại thời gian kết thúc.', 'error')
        return redirect(url_for('edit_assignment', assignment_id=a.id))
    if a.starts_at and a.due_at and a.starts_at >= a.due_at:
        flash('Chưa thể giao bài: thời gian kết thúc phải sau thời gian bắt đầu.', 'error')
        return redirect(url_for('edit_assignment', assignment_id=a.id))
    a.is_published = True
    db.session.commit()
    label, _ = assignment_status(a)
    if label == 'Chưa mở':
        flash('Đã giao bài. Học sinh sẽ thấy bài và có thể làm khi đến giờ bắt đầu.', 'ok')
    else:
        flash('Đã giao bài cho học sinh. Học sinh thuộc các lớp đã chọn có thể vào làm ngay.', 'ok')
    return redirect(request.referrer or url_for('edit_assignment', assignment_id=a.id))

@app.route('/teacher/assignment/<int:assignment_id>/unpublish', methods=['POST'])
def unpublish_assignment(assignment_id):
    if not teacher_only(): return redirect(url_for('login'))
    a = db.session.get(Assignment, assignment_id)
    if not a or a.created_by != me().id: return 'Không có quyền', 403
    a.is_published = False
    db.session.commit()
    flash('Đã thu hồi bài. Học sinh sẽ không còn nhìn thấy bài này.', 'ok')
    return redirect(request.referrer or url_for('edit_assignment', assignment_id=a.id))

@app.route('/teacher/assignment/<int:assignment_id>/remove/<int:qid>', methods=['POST'])
def remove_from_assignment(assignment_id, qid):
    if not teacher_only(): return redirect(url_for('login'))
    a = db.session.get(Assignment, assignment_id)
    if a and a.created_by == me().id:
        AssignmentQuestion.query.filter_by(assignment_id=a.id, question_id=qid).delete(); db.session.commit(); flash('Đã bỏ câu khỏi đề.', 'ok')
    return redirect(url_for('edit_assignment', assignment_id=assignment_id))

@app.route('/teacher/assignment/<int:assignment_id>/results')
def assignment_results(assignment_id):
    if not teacher_only(): return redirect(url_for('login'))
    a = db.session.get(Assignment, assignment_id)
    if not a or a.created_by != me().id: return 'Không có quyền', 403
    subs = Submission.query.filter_by(assignment_id=a.id, submitted=True).order_by(Submission.submitted_at.desc()).all()
    rows = [(s, db.session.get(User, s.student_id)) for s in subs]
    return render_template('results.html', assignment=a, rows=rows)

@app.route('/teacher/assignment/<int:assignment_id>/export.xlsx')
def export_results(assignment_id):
    if not teacher_only(): return redirect(url_for('login'))
    a = db.session.get(Assignment, assignment_id)
    if not a or a.created_by != me().id: return 'Không có quyền', 403
    subs = Submission.query.filter_by(assignment_id=a.id, submitted=True).all(); wb = Workbook(); ws = wb.active; ws.title = 'Bảng điểm'
    ws.append(['STT', 'Họ và tên', 'Lớp', 'Điểm trắc nghiệm', 'Điểm tự luận', 'Tổng điểm', 'Thang điểm', 'Thời gian nộp'])
    for cell in ws[1]:
        cell.font = Font(bold=True); cell.fill = PatternFill('solid', fgColor='D9EAF7'); cell.alignment = Alignment(horizontal='center')
    for i, s in enumerate(subs, 1):
        st = db.session.get(User, s.student_id); c = db.session.get(Classroom, st.classroom_id) if st and st.classroom_id else None
        ws.append([i, st.full_name if st else '', c.name if c else '', round(s.auto_score, 2), round(s.manual_score, 2), s.total_score,
                   s.max_score, s.submitted_at.strftime('%d/%m/%Y %H:%M') if s.submitted_at else ''])
    for col, w in {'A': 6, 'B': 28, 'C': 12, 'D': 18, 'E': 15, 'F': 12, 'G': 12, 'H': 20}.items(): ws.column_dimensions[col].width = w
    bio = BytesIO(); wb.save(bio); bio.seek(0); safe = re.sub(r'[^A-Za-z0-9_-]', '_', a.title)
    return send_file(bio, as_attachment=True, download_name=f'Bang_diem_{safe}.xlsx', mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')

@app.route('/teacher/analytics')
def analytics():
    if not teacher_only(): return redirect(url_for('login'))
    assignments = Assignment.query.filter_by(created_by=me().id).all(); rows = []
    for a in assignments:
        subs = Submission.query.filter_by(assignment_id=a.id, submitted=True).all()
        avg = round(sum(s.total_score for s in subs) / len(subs), 2) if subs else 0; mx = max([s.total_score for s in subs], default=0)
        rows.append((a, assignment_class_names(a), len(subs), avg, mx))
    return render_template('analytics.html', rows=rows)

@app.route('/teacher/submission/<int:sid>', methods=['GET', 'POST'])
def grade_submission(sid):
    if not teacher_only(): return redirect(url_for('login'))
    s = db.session.get(Submission, sid); a = db.session.get(Assignment, s.assignment_id) if s else None
    if not s or not a or a.created_by != me().id: return 'Không có quyền', 403
    answers = Answer.query.filter_by(submission_id=s.id).all(); qs = [db.session.get(BankQuestion, ans.question_id) for ans in answers]
    score_map = question_score_map(a, qs)
    if request.method == 'POST':
        manual = 0
        for ans, q in zip(answers, qs):
            if q and q.qtype == 'essay':
                qmax = score_map.get(q.id, 0)
                try: pts = max(0, min(qmax, float(request.form.get(f'score_{ans.id}', 0) or 0)))
                except: pts = 0
                ans.manual_score = pts; ans.teacher_note = request.form.get(f'note_{ans.id}', '').strip(); manual += pts
        s.manual_score = round(manual, 4); db.session.commit(); flash('Đã lưu điểm tự luận.', 'ok'); return redirect(url_for('assignment_results', assignment_id=a.id))
    return render_template('grade_submission.html', submission=s, assignment=a, student=db.session.get(User, s.student_id),
                           data=[(ans, q, score_map.get(q.id, 0)) for ans, q in zip(answers, qs)])

@app.route('/student')
def student_dashboard():
    if not student_only(): return redirect(url_for('login'))
    u = me(); c = db.session.get(Classroom, u.classroom_id) if u.classroom_id else None
    if u.classroom_id:
        ids = [x.assignment_id for x in AssignmentClassroom.query.filter_by(classroom_id=u.classroom_id).all()]
        assignments = Assignment.query.filter(Assignment.id.in_(ids), Assignment.is_published == True).order_by(Assignment.id.desc()).all() if ids else []
        lessons = Lesson.query.filter_by(classroom_id=u.classroom_id, is_published=True).order_by(Lesson.id.desc()).all()
    else:
        assignments = []; lessons = []
    subs = {s.assignment_id: s for s in Submission.query.filter_by(student_id=u.id).all()}
    return render_template('student_dashboard.html', classroom=c, assignments=assignments, lessons=lessons, subs=subs)

@app.route('/student/lesson/<int:lid>')
def student_lesson(lid):
    if not student_only(): return redirect(url_for('login'))
    x = db.session.get(Lesson, lid)
    if not x or x.classroom_id != me().classroom_id or not x.is_published: return 'Không có quyền', 403
    return render_template('student_lesson.html', lesson=x)

@app.route('/student/assignment/<int:assignment_id>/retake', methods=['POST'])
def retake_assignment(assignment_id):
    if not student_only(): return redirect(url_for('login'))
    a = db.session.get(Assignment, assignment_id); u = me(); s = Submission.query.filter_by(assignment_id=assignment_id, student_id=u.id).first()
    if not a or not s or not a.allow_retake or not student_has_assignment(a, u.classroom_id): return 'Không được phép làm lại', 403
    now = local_now()
    if a.starts_at and now < a.starts_at: flash('Bài chưa đến giờ mở.', 'error'); return redirect(url_for('student_dashboard'))
    if a.due_at and now > a.due_at: flash('Bài đã hết thời gian làm.', 'error'); return redirect(url_for('student_dashboard'))
    Answer.query.filter_by(submission_id=s.id).delete(); s.started_at = local_now(); s.submitted_at = None; s.submitted = False
    s.auto_score = 0; s.manual_score = 0; s.max_score = a.score_scale; db.session.commit(); return redirect(url_for('do_assignment', assignment_id=a.id))

@app.route('/student/assignment/<int:assignment_id>', methods=['GET', 'POST'])
def do_assignment(assignment_id):
    if not student_only(): return redirect(url_for('login'))
    u = me(); a = db.session.get(Assignment, assignment_id)
    if not a or not student_has_assignment(a, u.classroom_id) or not a.is_published: return 'Không có quyền', 403
    now = local_now(); sub = Submission.query.filter_by(assignment_id=a.id, student_id=u.id).first()
    if a.starts_at and now < a.starts_at:
        flash('Bài chưa mở. Thời gian bắt đầu: ' + a.starts_at.strftime('%d/%m/%Y %H:%M'), 'error'); return redirect(url_for('student_dashboard'))
    if a.due_at and now > a.due_at and not (sub and sub.submitted):
        flash('Bài này đã hết thời gian làm.', 'error'); return redirect(url_for('student_dashboard'))
    if sub and sub.submitted: return render_template('student_result.html', assignment=a, submission=sub)
    if not sub:
        sub = Submission(assignment_id=a.id, student_id=u.id, started_at=local_now(), max_score=a.score_scale)
        db.session.add(sub); db.session.commit()
    links = AssignmentQuestion.query.filter_by(assignment_id=a.id).order_by(AssignmentQuestion.order_no).all()
    questions = [db.session.get(BankQuestion, x.question_id) for x in links]
    if a.shuffle_questions: random.Random(sub.id).shuffle(questions)
    score_map = question_score_map(a, questions)
    elapsed = int((datetime.now() - sub.started_at).total_seconds()); remaining = max(0, a.duration_minutes * 60 - elapsed)
    if a.due_at: remaining = min(remaining, max(0, int((a.due_at - datetime.now()).total_seconds())))
    if request.method == 'POST' or remaining <= 0:
        Answer.query.filter_by(submission_id=sub.id).delete(); auto = 0
        for q in questions:
            val = request.form.get(f'q_{q.id}', '').strip(); pts = 0
            if q.qtype == 'mcq' and val.upper() == (q.correct_answer or '').upper():
                pts = score_map.get(q.id, 0); auto += pts
            elif q.qtype == 'short' and normalize_short_answer(val) == normalize_short_answer(q.correct_answer):
                pts = score_map.get(q.id, 0); auto += pts
            db.session.add(Answer(submission_id=sub.id, question_id=q.id, answer_text=val, auto_score=pts))
        sub.auto_score = round(auto, 4); sub.max_score = a.score_scale; sub.submitted = True; sub.submitted_at = datetime.now()
        db.session.commit(); session.clear()
        return render_template('submitted.html', score=round(auto, 2), max_score=a.score_scale, show_result=a.show_result)
    display = []
    for q in questions:
        opts = [('A', q.option_a), ('B', q.option_b), ('C', q.option_c), ('D', q.option_d)]
        if q.qtype == 'mcq' and a.shuffle_options: random.Random(sub.id * 100000 + q.id).shuffle(opts)
        display.append((q, opts, score_map.get(q.id, 0)))
    return render_template('do_assignment.html', assignment=a, display=display, remaining=remaining)

def ensure_v70_schema():
    """Nâng cơ sở dữ liệu cũ lên V7.0 mà không xóa lớp, học sinh hay bài làm."""
    try:
        insp = inspect(db.engine)
        cols = {c['name'] for c in insp.get_columns('bank_question')}
        if 'difficulty' not in cols:
            db.session.execute(text("ALTER TABLE bank_question ADD COLUMN difficulty VARCHAR(20) DEFAULT 'Trung bình'"))
        if 'domain' not in cols:
            db.session.execute(text("ALTER TABLE bank_question ADD COLUMN domain VARCHAR(20) DEFAULT 'Đại số'"))
            db.session.commit()
    except Exception as e:
        db.session.rollback()
        print('V7.0 schema migration warning:', e)


def classify_existing_builtin_questions():
    """Gán mức độ và phân môn cho gói NH1000 đã nạp ở các phiên bản trước."""
    rows = BankQuestion.query.filter(BankQuestion.explanation.like('§NH1000§%')).order_by(BankQuestion.id).all()
    by_grade_type = {}
    for q in rows:
        by_grade_type.setdefault((q.grade, q.qtype), []).append(q)
    # Mỗi khối: Dễ 100 (90 TN+10 ngắn), TB 100 (70 TN+20 ngắn+10 TL), Khó 50 (20 TN+10 ngắn+20 TL)
    splits = {
        'mcq': [('Dễ',90),('Trung bình',70),('Khó',20)],
        'short':[('Dễ',10),('Trung bình',20),('Khó',10)],
        'essay':[('Trung bình',10),('Khó',20)],
    }
    changed=0
    geo_terms=['hình học','góc','đường thẳng','tứ giác','tam giác','đường tròn','hệ thức lượng']
    for key, qs in by_grade_type.items():
        qtype=key[1]; pos=0
        for level,n in splits.get(qtype,[('Trung bình',len(qs))]):
            for q in qs[pos:pos+n]:
                if q.difficulty != level:
                    q.difficulty=level; changed+=1
            pos += n
        for q in qs[pos:]:
            if not q.difficulty:
                q.difficulty='Trung bình'; changed+=1
        # Phân môn được suy ra từ chủ đề, áp dụng cho toàn bộ câu trong nhóm.
        for q in qs:
            inferred='Hình học' if any(t in (q.topic or '').lower() for t in geo_terms) else 'Đại số'
            if getattr(q,'domain',None) != inferred:
                q.domain=inferred; changed+=1
    if changed: db.session.commit()



def run_v8_migrations():
    """
    Migration nhẹ, chạy an toàn nhiều lần.
    Không xóa dữ liệu cũ. Các bản sau có thể nối thêm migration mới theo số version.
    """
    try:
        db.session.execute(text(
            "CREATE TABLE IF NOT EXISTS app_schema_migration "
            "(version INTEGER PRIMARY KEY, applied_at VARCHAR(40) NOT NULL)"
        ))
        done = {int(r[0]) for r in db.session.execute(text(
            "SELECT version FROM app_schema_migration"
        )).fetchall()}

        migrations = []

        def migration_800():
            # Index phục vụ ngân hàng câu hỏi lớn + bài làm online.
            statements = [
                "CREATE INDEX IF NOT EXISTS ix_bank_question_owner_grade ON bank_question (owner_id, grade)",
                "CREATE INDEX IF NOT EXISTS ix_bank_question_filter ON bank_question (grade, domain, difficulty, qtype)",
                "CREATE INDEX IF NOT EXISTS ix_assignment_created_by ON assignment (created_by)",
                "CREATE INDEX IF NOT EXISTS ix_submission_assignment_student ON submission (assignment_id, student_id)",
                "CREATE INDEX IF NOT EXISTS ix_assignment_classroom_class ON assignment_classroom (classroom_id, assignment_id)"
            ]
            for stmt in statements:
                db.session.execute(text(stmt))

        migrations.append((800, migration_800))

        for version, fn in migrations:
            if version in done:
                continue
            fn()
            db.session.execute(
                text("INSERT INTO app_schema_migration(version, applied_at) VALUES (:v, :t)"),
                {"v": version, "t": local_now().isoformat(timespec='seconds')}
            )
            db.session.commit()
            print(f'Applied schema migration {version}')
    except Exception as e:
        db.session.rollback()
        print('V8 migration warning:', e)


def seed():
    if not User.query.filter_by(username='giaovien').first():
        db.session.add(User(username='giaovien', password_hash=generate_password_hash('123456'), full_name='Giáo viên Toán', role='teacher'))
    if not Classroom.query.first():
        db.session.add_all([Classroom(name='6A1', grade='6'), Classroom(name='7A1', grade='7'), Classroom(name='8A1', grade='8'), Classroom(name='9A1', grade='9')])
    db.session.commit(); t = User.query.filter_by(username='giaovien').first()
    if t and not SiteSetting.query.filter_by(owner_id=t.id).first():
        db.session.add(SiteSetting(owner_id=t.id, teacher_label=t.full_name)); db.session.commit()

with app.app_context():
    db.create_all(); ensure_v70_schema(); run_v8_migrations(); seed(); classify_existing_builtin_questions(); repair_k12_fraction_equations()

@app.route('/health')
def health():
    return {
        'status': 'ok',
        'version': '8.0-online',
        'timezone': APP_TIMEZONE,
        'database': 'postgresql' if str(app.config['SQLALCHEMY_DATABASE_URI']).startswith('postgresql') else 'sqlite'
    }, 200

@app.route('/ready')
def ready():
    try:
        db.session.execute(text('SELECT 1'))
        return {'status': 'ready', 'version': '8.0-online'}, 200
    except Exception as e:
        db.session.rollback()
        return {'status': 'not-ready', 'error': str(e)[:160]}, 503

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=int(os.environ.get('PORT', '5000')), debug=os.environ.get('FLASK_DEBUG', '0') == '1')
