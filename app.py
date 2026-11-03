"""
PDF Pro Converter - Backend API
Supports Arabic RTL, multi-language, and professional conversion
"""

import os
import json
import uuid
import logging
from logging.handlers import RotatingFileHandler
import psutil
from werkzeug.middleware.proxy_fix import ProxyFix
import gcs_helper
import hashlib
import hmac
import threading
from datetime import datetime, timedelta
from functools import wraps
import zipfile
import fitz  # PyMuPDF

from flask import Flask, request, jsonify, send_file, send_from_directory, render_template, make_response, g, redirect
from flask_cors import CORS
from flask_sqlalchemy import SQLAlchemy
from flask_jwt_extended import (
    JWTManager, create_access_token, jwt_required,
    get_jwt_identity, verify_jwt_in_request
)
from werkzeug.security import generate_password_hash, check_password_hash
from werkzeug.utils import secure_filename
from dotenv import load_dotenv
from flask_limiter import Limiter
from flask_limiter.util import get_remote_address
from flask_talisman import Talisman

from google.oauth2 import id_token
from google.auth.transport import requests as google_requests

from converter import PDFConverter

load_dotenv(os.path.join(os.path.dirname(os.path.abspath(__file__)), ".env"))

# ─────────────────────────────────────────────
# App Configuration
# ─────────────────────────────────────────────
app = Flask(__name__)
# Enable strict CORS for production and local development
# Enable strict CORS for production and local development
CORS(app, origins="*", expose_headers=["Content-Disposition"])
# Talisman disabled HTTPS redirect - Cloudflare already handles SSL termination
# Talisman was causing POST requests to fail (redirect 301 converts POST to GET)
Talisman(app, content_security_policy=None, force_https=False)

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
import tempfile

FRONTEND_DIR = os.path.join(os.path.dirname(BASE_DIR), 'android-app', 'www')
UPLOAD_FOLDER = os.path.join(tempfile.gettempdir(), 'pdfpro_uploads')
OUTPUT_FOLDER = os.path.join(tempfile.gettempdir(), 'pdfpro_outputs')
DB_PATH = os.path.join(BASE_DIR, 'database.db')

os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(OUTPUT_FOLDER, exist_ok=True)

db_url = os.getenv('DATABASE_URL')
if db_url:
    if db_url.startswith("postgres://"):
        db_url = db_url.replace("postgres://", "postgresql://", 1)
    app.config['SQLALCHEMY_DATABASE_URI'] = db_url
else:
    app.config['SQLALCHEMY_DATABASE_URI'] = f'sqlite:///{DB_PATH}'
app.config['SQLALCHEMY_TRACK_MODIFICATIONS'] = False
# Basic Rate Limiter setup (using in-memory storage for simplicity)

def get_cf_ip():
    return request.headers.get('CF-Connecting-IP', request.headers.get('X-Forwarded-For', request.remote_addr))

def get_system_stats():
    """Get real CPU and RAM usage using psutil (cross-platform)"""
    try:
        cpu_percent = round(psutil.cpu_percent(interval=0.5), 1)
        mem = psutil.virtual_memory()
        ram_percent = round(mem.percent, 1)
        return cpu_percent, ram_percent
    except Exception:
        return 0.0, 0.0

def send_conversion_email(recipient_email, download_url, filename):
    import smtplib
    from email.mime.text import MIMEText
    from email.mime.multipart import MIMEMultipart

    smtp_server = os.getenv('SMTP_SERVER', '').strip()
    smtp_port = os.getenv('SMTP_PORT', '587').strip()
    smtp_user = os.getenv('SMTP_USER', '').strip()
    smtp_pass = os.getenv('SMTP_PASSWORD', '').strip()
    smtp_from = os.getenv('SMTP_FROM_EMAIL', 'PDF Pro <noreply@hibdf.com>').strip()

    if not smtp_server or not smtp_user or not smtp_pass:
        logger.warning(f"SMTP not configured. Skipping email to {recipient_email}. Link: {download_url}")
        return False

    try:
        msg = MIMEMultipart('alternative')
        msg['Subject'] = '📄 ملفك المحول جاهز للتحميل! - PDF Pro'
        msg['From'] = smtp_from
        msg['To'] = recipient_email

        html = f"""
        <html>
        <head>
            <style>
                body {{ font-family: 'Cairo', Arial, sans-serif; direction: rtl; text-align: right; background-color: #f4f5f7; color: #333; }}
                .container {{ max-width: 600px; margin: 30px auto; padding: 25px; background: #ffffff; border-radius: 12px; border: 1px solid #e1e4e8; box-shadow: 0 4px 6px rgba(0,0,0,0.05); }}
                .logo {{ font-size: 28px; font-weight: bold; color: #a855f7; text-align: center; margin-bottom: 20px; }}
                .title {{ font-size: 20px; font-weight: bold; color: #2d3748; margin-bottom: 15px; text-align: center; }}
                .desc {{ font-size: 15px; color: #4a5568; line-height: 1.6; margin-bottom: 25px; text-align: center; }}
                .btn-container {{ text-align: center; margin-bottom: 30px; }}
                .btn {{ display: inline-block; padding: 12px 30px; background: #a855f7; color: #ffffff !important; text-decoration: none; font-weight: bold; border-radius: 8px; font-size: 16px; box-shadow: 0 4px 6px rgba(168, 85, 247, 0.2); }}
                .footer {{ font-size: 12px; color: #718096; text-align: center; border-top: 1px solid #edf2f7; padding-top: 15px; margin-top: 25px; }}
            </style>
        </head>
        <body>
            <div class="container">
                <div class="logo">📄 PDF Pro</div>
                <div class="title">تهانينا! تم تحويل ملفك بنجاح 🎉</div>
                <div class="desc">
                    لقد قمنا بمعالجة وتحويل مستندك <strong>{filename}</strong> بنجاح باستخدام أفضل محركات التنسيق الذكية.<br>
                    الرابط متاح للتحميل لمدة <strong>يومين فقط (48 ساعة)</strong>، وسيتم مسحه بعدها تلقائياً لحماية خصوصيتك.
                </div>
                <div class="btn-container">
                    <a href="{download_url}" class="btn">📥 تحميل ملف Word الآن</a>
                </div>
                <div class="footer">
                    هذه الرسالة مرسلة تلقائياً من نظام PDF Pro Converter.<br>
                    يرجى عدم الرد على هذا البريد الإلكتروني.
                </div>
            </div>
        </body>
        </html>
        """
        msg.attach(MIMEText(html, 'html', 'utf-8'))

        server = smtplib.SMTP(smtp_server, int(smtp_port))
        server.starttls()
        server.login(smtp_user, smtp_pass)
        server.sendmail(smtp_from, [recipient_email], msg.as_string())
        server.quit()
        logger.info(f"Successfully sent conversion email to {recipient_email}")
        return True
    except Exception as e:
        logger.error(f"Failed to send email to {recipient_email}: {e}")
        return False

def start_queue_scheduler():
    import time
    def scheduler_loop():
        with app.app_context():
            logger.info("🚀 Background Queue Scheduler Thread Started")
            while True:
                try:
                    time.sleep(5)
                    # Mark stuck processing tasks as failed (older than 2 hours)
                    two_hours_ago = datetime.utcnow() - timedelta(hours=2)
                    stuck_tasks = Conversion.query.filter(Conversion.status == 'processing', Conversion.created_at < two_hours_ago).all()
                    for t in stuck_tasks:
                        t.status = 'failed'
                        t.error_message = 'Timeout: Process took too long'
                    if stuck_tasks:
                        db.session.commit()

                    # Process queued items
                    active_count = Conversion.query.filter_by(status='processing').count()
                    limit_setting = AppSetting.query.filter_by(key='max_concurrent_conversions').first()
                    max_concurrent = int(limit_setting.value if limit_setting else '3')
                    
                    if active_count < max_concurrent:
                        query = Conversion.query.filter_by(status='queued').order_by(Conversion.created_at.asc())
                        if app.config['SQLALCHEMY_DATABASE_URI'].startswith('postgresql'):
                            next_conv = query.with_for_update(skip_locked=True).first()
                        else:
                            next_conv = query.first()
                        if next_conv:
                            if next_conv.params:
                                try:
                                    params = json.loads(next_conv.params)
                                except Exception:
                                    logger.error(f"Failed to parse params for {next_conv.conversion_id}")
                                    next_conv.status = 'failed'
                                    next_conv.error_message = "خطأ في قراءة بارامترات التحويل"
                                    db.session.commit()
                                    continue
                                
                                next_conv.status = 'processing'
                                db.session.commit()
                                
                                thread = threading.Thread(
                                    target=run_async_conversion,
                                    args=(
                                        next_conv.conversion_id,
                                        params['input_path'],
                                        params['output_path'],
                                        params['language'],
                                        params['gemini_api_key'],
                                        next_conv.user_id,
                                        params.get('original_image_path'),
                                        params['conversion_mode'],
                                        params.get('gcs_input_name')
                                    )
                                )
                                thread.start()
                                logger.info(f"Queue Scheduler: Started queued conversion {next_conv.conversion_id}")
                except Exception as e:
                    logger.error(f"Queue Scheduler loop error: {e}")
    t = threading.Thread(target=scheduler_loop, daemon=True)
    t.start()

def start_file_cleanup_scheduler():
    import time
    def cleanup_loop():
        with app.app_context():
            logger.info("🚀 Background File Cleanup Thread Started")
            while True:
                try:
                    time.sleep(3600)
                    now = datetime.utcnow()
                    expiration_limit = now - timedelta(days=2)
                    
                    cleaned_count = 0
                    old_conversions = Conversion.query.filter(Conversion.created_at <= expiration_limit).all()
                    for conv in old_conversions:
                        for fname in os.listdir(OUTPUT_FOLDER):
                            if fname.startswith(conv.conversion_id):
                                fpath = os.path.join(OUTPUT_FOLDER, fname)
                                try:
                                    os.remove(fpath)
                                    cleaned_count += 1
                                except Exception:
                                    pass
                        for fname in os.listdir(UPLOAD_FOLDER):
                            if fname.startswith(conv.conversion_id):
                                fpath = os.path.join(UPLOAD_FOLDER, fname)
                                try:
                                    os.remove(fpath)
                                    cleaned_count += 1
                                except Exception:
                                    pass
                                    
                    if cleaned_count > 0:
                        logger.info(f"Cleanup Scheduler: Cleaned up {cleaned_count} expired files.")
                except Exception as e:
                    logger.error(f"Cleanup Scheduler loop error: {e}")
    t = threading.Thread(target=cleanup_loop, daemon=True)
    t.start()

limiter = Limiter(
    get_cf_ip,
    app=app,
    default_limits=["10000 per day", "5000 per hour"],
    storage_uri="memory://"
)

# Apply basic security headers (Talisman)
# Disable HTTPS enforcing in dev if needed, but assuming prod is behind Cloudflare
# talisman = Talisman(app, content_security_policy=None, force_https=False)

app.config['JWT_SECRET_KEY'] = os.getenv('JWT_SECRET', 'pdf-pro-secret-2026-ultra-secure')
app.config['JWT_ACCESS_TOKEN_EXPIRES'] = timedelta(days=7)
app.config['MAX_CONTENT_LENGTH'] = 500 * 1024 * 1024  # 500MB

db = SQLAlchemy(app)
jwt = JWTManager(app)
converter = PDFConverter()

logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

ALLOWED_EXTENSIONS = {'pdf'}

PDF_MAGIC_BYTES = b'%PDF'
MAX_FILE_SIZE = 200 * 1024 * 1024  # 200MB per file

def validate_pdf_file(file_path):
    """Validate file is a real PDF by checking magic bytes signature."""
    try:
        with open(file_path, 'rb') as f:
            header = f.read(4)
        return header == PDF_MAGIC_BYTES
    except Exception:
        return False

def sanitize_input(value, max_length=255):
    """Basic input sanitization - strip whitespace and limit length."""
    if value is None:
        return value
    value = str(value).strip()
    return value[:max_length]

def validate_email(email):
    """Validate email format."""
    import re
    if not email:
        return False
    pattern = r'^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$'
    return bool(re.match(pattern, email)) and len(email) <= 255

def validate_uuid(value):
    """Validate that a value is a proper UUID format to prevent injection."""
    import re
    if not value:
        return False
    pattern = r'^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$'
    return bool(re.match(pattern, str(value).lower()))

# ─────────────────────────────────────────────
# Database Models
# ─────────────────────────────────────────────

class AdminUser(db.Model):
    __tablename__ = 'admin_users'
    id = db.Column(db.Integer, primary_key=True)
    username = db.Column(db.String(80), unique=True, nullable=False)
    email = db.Column(db.String(120), unique=True, nullable=False)
    password_hash = db.Column(db.String(256), nullable=False)
    role = db.Column(db.String(20), default='admin')  # superadmin / admin
    created_at = db.Column(db.DateTime, default=datetime.utcnow)
    last_login = db.Column(db.DateTime)

    def set_password(self, password):
        self.password_hash = generate_password_hash(password)

    def check_password(self, password):
        return check_password_hash(self.password_hash, password)


class AppUser(db.Model):
    __tablename__ = 'app_users'
    id = db.Column(db.Integer, primary_key=True)
    # Auth fields
    google_id = db.Column(db.String(255), unique=True, nullable=True)
    email = db.Column(db.String(255), unique=True, nullable=True)
    device_id = db.Column(db.String(100), unique=True, nullable=True)  # Made nullable for new accounts
    
    # Billing & Subscription fields
    plan = db.Column(db.String(20), default='free')  # free / premium
    subscription_expiry = db.Column(db.DateTime, nullable=True)
    play_purchase_token = db.Column(db.String(500), nullable=True)
    auto_renew = db.Column(db.Boolean, default=True)
    paddle_subscription_id = db.Column(db.String(100), nullable=True)
    paddle_cancel_url = db.Column(db.String(500), nullable=True)
    fcm_token = db.Column(db.String(500), nullable=True)
    
    # User Preferences & Stats
    platform = db.Column(db.String(20), default='android')
    language = db.Column(db.String(10), default='ar')
    conversions_today = db.Column(db.Integer, default=0)
    total_conversions = db.Column(db.Integer, default=0)
    bonus_conversions = db.Column(db.Integer, default=0)
    last_conversion = db.Column(db.DateTime)
    created_at = db.Column(db.DateTime, default=datetime.utcnow)
    last_seen = db.Column(db.DateTime, default=datetime.utcnow)
    country = db.Column(db.String(50))
    ip_address = db.Column(db.String(100))
    fingerprint = db.Column(db.String(255))


class Conversion(db.Model):
    __tablename__ = 'conversions'
    id = db.Column(db.Integer, primary_key=True)
    conversion_id = db.Column(db.String(100), unique=True, nullable=False)
    user_id = db.Column(db.Integer, db.ForeignKey('app_users.id'), nullable=True)
    device_id = db.Column(db.String(100), nullable=True)
    filename = db.Column(db.String(255))
    original_size = db.Column(db.Integer)
    output_size = db.Column(db.Integer)
    pages = db.Column(db.Integer)
    language_detected = db.Column(db.String(20))
    status = db.Column(db.String(20), default='pending')  # pending/queued/processing/done/failed
    processing_time = db.Column(db.Float)
    created_at = db.Column(db.DateTime, default=datetime.utcnow)
    completed_at = db.Column(db.DateTime)
    error_message = db.Column(db.Text)
    params = db.Column(db.Text, nullable=True)
    email = db.Column(db.String(255), nullable=True)

class PaymentTransaction(db.Model):
    __tablename__ = 'payment_transactions'
    id = db.Column(db.Integer, primary_key=True)
    user_id = db.Column(db.Integer, db.ForeignKey('app_users.id'), nullable=True)
    paddle_transaction_id = db.Column(db.String(100), unique=True, nullable=False)
    amount = db.Column(db.Float, nullable=False)
    currency = db.Column(db.String(10), default='USD')
    status = db.Column(db.String(50), default='completed')
    country = db.Column(db.String(50))
    created_at = db.Column(db.DateTime, default=datetime.utcnow)


class Advertisement(db.Model):
    __tablename__ = 'advertisements'
    id = db.Column(db.Integer, primary_key=True)
    name = db.Column(db.String(200), nullable=False)
    ad_type = db.Column(db.String(50))  # banner / interstitial / rewarded / video
    ad_unit_id = db.Column(db.String(200))
    position = db.Column(db.String(100), default='top')  # top / bottom / middle / after_upload / popup
    video_url = db.Column(db.String(500), nullable=True)  # رابط الفيديو الإعلاني
    ad_link = db.Column(db.String(500), nullable=True)    # رابط توجيه الزائر عند النقر
    ad_code = db.Column(db.Text, nullable=True)           # كود إعلانات جوجل أو HTML
    is_active = db.Column(db.Boolean, default=True)
    impressions = db.Column(db.Integer, default=0)
    clicks = db.Column(db.Integer, default=0)
    revenue = db.Column(db.Float, default=0.0)
    created_at = db.Column(db.DateTime, default=datetime.utcnow)
    updated_at = db.Column(db.DateTime, default=datetime.utcnow)


class PageVisit(db.Model):
    __tablename__ = 'page_visits'
    id = db.Column(db.Integer, primary_key=True)
    device_id = db.Column(db.String(100))
    ip_address = db.Column(db.String(100))
    page = db.Column(db.String(100))
    duration_seconds = db.Column(db.Integer)
    country = db.Column(db.String(50))
    created_at = db.Column(db.DateTime, default=datetime.utcnow)


class AppSetting(db.Model):
    __tablename__ = 'app_settings'
    id = db.Column(db.Integer, primary_key=True)
    key = db.Column(db.String(100), unique=True, nullable=False)
    value = db.Column(db.Text)
    description = db.Column(db.Text)


# ─────────────────────────────────────────────
# Helpers
# ─────────────────────────────────────────────

def allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS


def admin_required(f):
    @wraps(f)
    def decorated(*args, **kwargs):
        try:
            verify_jwt_in_request()
            identity = get_jwt_identity()
            admin = AdminUser.query.filter_by(username=identity).first()
            if not admin:
                return jsonify({'error': 'Admin access required'}), 403
        except Exception:
            return jsonify({'error': 'Unauthorized'}), 401
        return f(*args, **kwargs)
    return decorated


def get_or_create_user(device_id, fingerprint=None, ip_address=None, language='ar', country=None):
    # Try to find by fingerprint first (if provided) to catch incognito users
    user = None
    if fingerprint:
        user = AppUser.query.filter(
            db.or_(AppUser.device_id == device_id, AppUser.fingerprint == fingerprint)
        ).first()
    else:
        user = AppUser.query.filter_by(device_id=device_id).first()
        
    if not user:
        from sqlalchemy.exc import IntegrityError
        import uuid
        try:
            user = AppUser(device_id=device_id, language=language, fingerprint=fingerprint, ip_address=ip_address, country=country)
            db.session.add(user)
            db.session.commit()
        except IntegrityError:
            db.session.rollback()
            # It was likely created by a concurrent request, fetch it again
            user = AppUser.query.filter_by(device_id=device_id).first()
            if not user:
                # If still not found, generate a new device_id to avoid constraint failure
                new_dev_id = "dev_conflict_" + str(uuid.uuid4())[:12]
                user = AppUser(device_id=new_dev_id, language=language, fingerprint=fingerprint, ip_address=ip_address)
                db.session.add(user)
                db.session.commit()
    else:
        user.last_seen = datetime.utcnow()
        if fingerprint and not user.fingerprint:
            user.fingerprint = fingerprint
        if ip_address:
            user.ip_address = ip_address
        if country:
            user.country = country
        db.session.commit()
    return user


def get_setting(key, default=None):
    setting = AppSetting.query.filter_by(key=key).first()
    return setting.value if setting else default


# ─────────────────────────────────────────────
# Frontend Static File Serving
# ─────────────────────────────────────────────

@app.route('/')
def serve_index():
    return send_from_directory(FRONTEND_DIR, 'index.html')

@app.route('/<path:filename>')
def serve_static(filename):
    # Prevent path traversal - ensure the resolved path stays within FRONTEND_DIR
    safe_path = os.path.normpath(os.path.join(FRONTEND_DIR, filename))
    if not safe_path.startswith(os.path.abspath(FRONTEND_DIR)):
        return jsonify({'error': 'Not Found'}), 404
    return send_from_directory(FRONTEND_DIR, filename)

@app.after_request
def add_header(response):
    # Disable caching for index.html and other static html/js files served
    if request.path == '/' or request.path.endswith('.html') or request.path.endswith('.js'):
        response.headers['Cache-Control'] = 'no-store, no-cache, must-revalidate, post-check=0, pre-check=0, max-age=0'
        response.headers['Pragma'] = 'no-cache'
        response.headers['Expires'] = '-1'
    
    # Security Headers
    response.headers['X-Content-Type-Options'] = 'nosniff'
    response.headers['X-Frame-Options'] = 'SAMEORIGIN'
    response.headers['X-XSS-Protection'] = '1; mode=block'
    response.headers['Referrer-Policy'] = 'strict-origin-when-cross-origin'
    response.headers['Permissions-Policy'] = 'camera=(), microphone=(), geolocation=()'
    response.headers['Cross-Origin-Resource-Policy'] = 'cross-origin'
    response.headers['Cross-Origin-Opener-Policy'] = 'same-origin-allow-popups'
    # Only add HSTS if behind HTTPS (Cloudflare handles SSL)
    response.headers['Strict-Transport-Security'] = 'max-age=31536000; includeSubDomains'
    # Remove server info disclosure
    response.headers.pop('Server', None)
    response.headers.pop('X-Powered-By', None)
    return response

# ─────────────────────────────────────────────
# Auth Routes
# ─────────────────────────────────────────────

@app.route('/api/admin/login', methods=['POST'])
@limiter.limit("10 per minute")
def admin_login():
    data = request.get_json()
    username = data.get('username', '')
    password = data.get('password', '')

    admin = AdminUser.query.filter_by(username=username).first()
    if not admin or not admin.check_password(password):
        return jsonify({'error': 'Invalid credentials'}), 401

    admin.last_login = datetime.utcnow()
    db.session.commit()

    token = create_access_token(identity=username)
    return jsonify({
        'access_token': token,
        'username': admin.username,
        'email': admin.email,
        'role': admin.role
    })


@app.route('/api/admin/me', methods=['GET'])
@admin_required
def admin_me():
    username = get_jwt_identity()
    admin = AdminUser.query.filter_by(username=username).first()
    return jsonify({
        'username': admin.username,
        'email': admin.email,
        'role': admin.role,
        'created_at': admin.created_at.isoformat()
    })


# ─────────────────────────────────────────────
# Conversion Routes (Mobile App)
# ─────────────────────────────────────────────

def run_async_conversion(conv_id, input_path, output_path, language, gemini_api_key, user_id, original_image_path=None, conversion_mode='auto', gcs_input_name=None):
    with app.app_context():
        logger.info(f"Starting async conversion for {conv_id}. Mode: {conversion_mode}, Has API Key: {bool(gemini_api_key)}, Lang: {language}")
        conv_record = Conversion.query.filter_by(conversion_id=conv_id).first()
        user = AppUser.query.get(user_id)
        if not conv_record or not user:
            return
            
        if gcs_input_name and os.getenv('GCS_BUCKET_NAME'):
            import gcs_helper
            client = gcs_helper.get_gcs_client()
            if client:
                bucket = client.bucket(os.getenv('GCS_BUCKET_NAME'))
                blob = bucket.blob(gcs_input_name)
                if blob.exists():
                    blob.download_to_filename(input_path)
        
        start_time = datetime.utcnow()
        try:
            def progress_callback(page_num, total_pages=None):
                if total_pages is None:
                    conv_record.error_message = str(page_num)
                else:
                    conv_record.error_message = f"{page_num}/{total_pages}"
                db.session.commit()
            
            result = converter.convert(
                input_path, output_path,
                language=language,
                mode=conversion_mode,
                gemini_api_key=gemini_api_key,
                progress_callback=progress_callback
            )
            
            processing_time = (datetime.utcnow() - start_time).total_seconds()
            
            conv_record.status = 'done'
            conv_record.output_size = os.path.getsize(output_path) if os.path.exists(output_path) else 0
            conv_record.pages = result.get('pages', 0)
            conv_record.language_detected = result.get('language', language)
            conv_record.processing_time = processing_time
            conv_record.completed_at = datetime.utcnow()
            conv_record.error_message = result.get('engine', 'local')
            
            today = datetime.utcnow().date()
            if not user.last_conversion or user.last_conversion.date() != today:
                user.conversions_today = 1
            else:
                user.conversions_today += 1
            user.total_conversions += 1
            user.last_conversion = datetime.utcnow()
            db.session.commit()
            
            if os.getenv('GCS_BUCKET_NAME'):
                import gcs_helper
                for fname in os.listdir(OUTPUT_FOLDER):
                    if fname.startswith(conv_id):
                        local_file = os.path.join(OUTPUT_FOLDER, fname)
                        gcs_helper.upload_to_gcs(local_file, f"outputs/{fname}")
                        try:
                            os.remove(local_file)
                        except Exception:
                            pass

            # Send email notification ONLY if the task was queued and user provided their email
            if conv_record.email and conv_record.params:
                site_url = os.getenv('SITE_URL', 'https://backend.hibdf.com').rstrip('/')
                download_url = f"{site_url}/api/download/{conv_id}"
                email_thread = threading.Thread(
                    target=send_conversion_email,
                    args=(conv_record.email, download_url, conv_record.filename)
                )
                email_thread.start()
            
        except Exception as e:
            logger.error(f"Async conversion failed for {conv_id}: {e}", exc_info=True)
            conv_record.status = 'failed'
            conv_record.error_message = str(e)
            db.session.commit()
        finally:
            try:
                os.remove(input_path)
            except Exception:
                pass
            # Clean up original image file if it was converted from an image
            if original_image_path:
                try:
                    os.remove(original_image_path)
                except Exception:
                    pass
@app.errorhandler(413)
def request_entity_too_large(error):
    return jsonify({'error': 'الملف كبير جداً (أكبر من 200 ميجا). يرجى التوجه لصفحة "تقسيم" لتقسيم الملف إلى أجزاء صغيرة قبل تحويله.'}), 413

# ─────────────────────────────────────────────
# Auth Routes (Google Sign-In)
# ─────────────────────────────────────────────

@app.route('/api/auth/google', methods=['POST'])
@limiter.limit("10 per minute")
def google_auth():
    data = request.get_json()
    token = data.get('idToken')
    device_id = request.headers.get('X-Device-ID', 'unknown')
    fingerprint = request.headers.get('X-Fingerprint')
    ip_address = request.headers.get('CF-Connecting-IP', request.headers.get('X-Forwarded-For', request.remote_addr))
    language = request.headers.get('X-Language', 'ar')
    
    if not token:
        return jsonify({'error': 'No ID token provided'}), 400
        
    try:
        # Verify the token with Google
        client_id = os.getenv('GOOGLE_CLIENT_ID', '867153923921-6a62f5i5sdjdsl402ml65gf9je96jkjs.apps.googleusercontent.com')
        idinfo = id_token.verify_oauth2_token(token, google_requests.Request(), client_id, clock_skew_in_seconds=600)
        
        google_id = idinfo['sub']
        email = idinfo.get('email', '')
        
        if not email.lower().endswith('@gmail.com'):
            return jsonify({'error': 'invalid_email_domain', 'message': 'عذراً، التطبيق يقبل حسابات @gmail.com فقط.'}), 403
        
        # First, find if the user already exists by google_id or email
        user = AppUser.query.filter((AppUser.google_id == google_id) | (AppUser.email == email)).first()

        # Check how many accounts this IP/Fingerprint has registered
        existing_accounts_count = 0
        if fingerprint or ip_address:
            conditions = []
            if fingerprint: conditions.append(AppUser.fingerprint == fingerprint)
            if ip_address: conditions.append(AppUser.ip_address == ip_address)
            existing_accounts_count = AppUser.query.filter(
                db.or_(*conditions), 
                AppUser.google_id.isnot(None)
            ).count()

        if user:
            # Existing Google user logs in on this device
            # Clear device_id from any other user to prevent UNIQUE constraint failure
            other_user = AppUser.query.filter(AppUser.device_id == device_id, AppUser.id != user.id).first()
            if other_user:
                other_user.device_id = None
                db.session.flush()
                
            user.device_id = device_id
            user.last_seen = datetime.utcnow()
            if fingerprint and not user.fingerprint: user.fingerprint = fingerprint
            if ip_address: user.ip_address = ip_address
        else:
            # This is a new Google account! 
            # Does the current device_id belong to an anonymous user?
            anon_user = None
            if fingerprint:
                anon_user = AppUser.query.filter(
                    db.or_(AppUser.device_id == device_id, AppUser.fingerprint == fingerprint),
                    AppUser.email == None
                ).first()
            else:
                anon_user = AppUser.query.filter_by(device_id=device_id, email=None).first()
            
            if anon_user:
                # Upgrade anonymous user to registered user
                user = anon_user
                user.email = email
                user.google_id = google_id
                user.last_seen = datetime.utcnow()
                # Apply bonus only if less than 2 accounts
                if existing_accounts_count >= 2:
                    user.bonus_conversions = 0
                else:
                    user.bonus_conversions = int(get_setting('login_bonus_conversions', '2'))
            else:
                # Clear device_id from any other registered user to prevent UNIQUE constraint failure
                other_user = AppUser.query.filter(AppUser.device_id == device_id).first()
                if other_user:
                    other_user.device_id = None
                    db.session.flush()
                    
                # Create a brand new user row
                user = AppUser(
                    google_id=google_id,
                    email=email,
                    device_id=device_id,
                    language=language,
                    fingerprint=fingerprint,
                    ip_address=ip_address
                )
                db.session.add(user)
                if existing_accounts_count >= 2:
                    user.bonus_conversions = 0
                else:
                    user.bonus_conversions = int(get_setting('login_bonus_conversions', '2'))

        db.session.commit()
        
        # Generate our own JWT for API access
        access_token = create_access_token(identity=str(user.id))
        
        return jsonify({
            'success': True,
            'token': access_token,
            'user': {
                'id': user.id,
                'email': user.email,
                'plan': user.plan,
                'subscription_expiry': user.subscription_expiry.isoformat() if user.subscription_expiry else None
            }
        })
        
    except ValueError as e:
        logger.error(f"Invalid Google token: {e}")
        return jsonify({'error': 'Invalid token'}), 401

# ─────────────────────────────────────────────
# Billing Routes (Google Play Validation)
# ─────────────────────────────────────────────

@app.route('/api/billing/verify', methods=['POST'])
@jwt_required()
def verify_billing():
    user_id = get_jwt_identity()
    user = AppUser.query.get(user_id)
    
    if not user:
        return jsonify({'error': 'User not found'}), 404
        
    data = request.get_json()
    purchase_token = data.get('purchaseToken')
    subscription_id = data.get('subscriptionId')
    package_name = data.get('packageName', 'com.pdfpro.converter')
    
    if not all([purchase_token, subscription_id]):
        return jsonify({'error': 'Missing billing details'}), 400
        
    try:
        # TODO: Implement real Server-Side Verification using google-api-python-client
        # This requires a Service Account JSON file from Google Play Console.
        # Example Implementation:
        # from google.oauth2 import service_account
        # from googleapiclient.discovery import build
        # creds = service_account.Credentials.from_service_account_file('play-service-account.json')
        # service = build('androidpublisher', 'v3', credentials=creds)
        # res = service.purchases().subscriptionsv2().get(packageName=package_name, token=purchase_token).execute()
        
        # BILLING NOT YET CONFIGURED - Reject all mock purchases
        # Real Google Play verification must be set up before accepting payments
        return jsonify({
            'error': 'نظام الدفع قيد الإعداد حالياً. سيتم تفعيله قريباً.',
            'message': 'Payment system is being configured. Coming soon.'
        }), 503
        
    except Exception as e:
        logger.error(f"Billing verification failed: {e}")
        return jsonify({'error': 'Verification failed, invalid receipt'}), 500

@app.route('/api/upload-chunk', methods=['POST'])
@limiter.limit("200 per hour")
def upload_chunk():
    try:
        chunk_index = request.form.get('chunk_index', type=int)
        total_chunks = request.form.get('total_chunks', type=int)
        file_id = request.form.get('file_id')
        
        if 'chunk' not in request.files or not file_id:
            return jsonify({'error': 'Missing chunk or file_id'}), 400
            
        if not validate_uuid(file_id):
            return jsonify({'error': 'Invalid file_id format'}), 400
        chunk_file = request.files['chunk']
        temp_path = os.path.join(UPLOAD_FOLDER, f"temp_{file_id}")
        
        # Append chunk to file
        with open(temp_path, 'ab') as f:
            f.write(chunk_file.read())
            
        return jsonify({'success': True, 'chunk_index': chunk_index})
    except Exception as e:
        logger.error(f"Error in upload_chunk: {e}")
        return jsonify({'error': str(e)}), 500

@app.route('/api/convert', methods=['POST'])
@limiter.limit("20 per hour; 5 per minute")
def convert_pdf():
    device_id = request.headers.get('X-Device-ID', 'unknown')
    language = request.headers.get('X-Language', 'ar')
    conversion_mode = request.headers.get('X-Mode', 'auto')
    gemini_api_key = request.headers.get('X-Gemini-API-Key', '').strip()
    
    # Fallback: use server-side env var if no key was sent from the client
    if not gemini_api_key:
        gemini_api_key = os.getenv('GEMINI_API_KEY', '').strip()
        if gemini_api_key:
            logger.info("No API key from client — using server-side GEMINI_API_KEY env var (keys will be rotated per page).")
        else:
            logger.warning("No Gemini API key provided (neither from client nor from GEMINI_API_KEY env var). Will use local engine.")

    is_logged_in = False
    try:
        verify_jwt_in_request(optional=True)
        user_id = get_jwt_identity()
        if user_id:
            user = AppUser.query.get(user_id)
            if user:
                is_logged_in = True
    except:
        pass
        
    if not is_logged_in:
        fingerprint = request.headers.get('X-Fingerprint')
        ip_address = request.headers.get('CF-Connecting-IP', request.headers.get('X-Forwarded-For', request.remote_addr))
        user = get_or_create_user(device_id, fingerprint, ip_address, language, request.headers.get('CF-IPCountry'))

    # Check lifetime limit for free users
    free_limit = int(get_setting('free_lifetime_limit', '2'))
    if is_logged_in:
        free_limit += 2
        
    try:
        bonus = user.bonus_conversions if getattr(user, 'bonus_conversions', None) is not None else 0
        if is_logged_in:
            free_limit += bonus
    except:
        pass

    effective_plan = user.plan if is_logged_in else 'free'

    if effective_plan == 'free':
        if user.total_conversions >= free_limit:
            return jsonify({
                'error': 'lifetime_limit_reached',
                'message': f'لقد استنفدت رصيدك المجاني ({free_limit} تحويلات). يرجى تسجيل الدخول أو الاشتراك للاستمرار.',
                'requires_purchase': True
            }), 403

    if 'file_id' in request.form:
        file_id = request.form.get('file_id', '')
        file_filename = request.form.get('filename', 'document.pdf')
        
        # Validate file_id is a proper UUID to prevent Path Traversal
        if not validate_uuid(file_id):
            return jsonify({'error': 'Invalid request'}), 400
        
        # Sanitize filename
        file_filename = sanitize_input(file_filename, 255)
        
        temp_path = os.path.join(UPLOAD_FOLDER, f"temp_{file_id}")
        # Ensure path stays within UPLOAD_FOLDER (path traversal prevention)
        if not os.path.abspath(temp_path).startswith(os.path.abspath(UPLOAD_FOLDER)):
            return jsonify({'error': 'Invalid request'}), 400
            
        if not os.path.exists(temp_path):
            return jsonify({'error': 'Uploaded file not found. It may have expired.'}), 400
            
        if not allowed_file(file_filename):
            try: os.remove(temp_path)
            except: pass
            return jsonify({'error': 'Invalid file. Only PDF files are allowed.'}), 400
            
        conv_id = str(uuid.uuid4())
        safe_name = f"{conv_id}.pdf"  # Rename to UUID-based name for security
        input_path = os.path.join(UPLOAD_FOLDER, safe_name)
        os.rename(temp_path, input_path)
        
        # Validate it's a real PDF by checking magic bytes
        if not validate_pdf_file(input_path):
            try: os.remove(input_path)
            except: pass
            return jsonify({'error': 'Invalid PDF file content.'}), 400
    else:
        if 'file' not in request.files:
            return jsonify({'error': 'No file provided'}), 400

        file = request.files['file']
        if not file or file.filename == '' or not allowed_file(file.filename):
            return jsonify({'error': 'Invalid file. Only PDF files are allowed.'}), 400

        conv_id = str(uuid.uuid4())
        safe_name = f"{conv_id}.pdf"  # Rename to UUID-based name for security
        input_path = os.path.join(UPLOAD_FOLDER, safe_name)
        file.save(input_path)
        
        # Validate it's a real PDF by checking magic bytes
        if not validate_pdf_file(input_path):
            try: os.remove(input_path)
            except: pass
            return jsonify({'error': 'Invalid PDF file content.'}), 400

    # Slice PDF if page range headers are provided
    page_from = request.headers.get('X-Page-From')
    page_to = request.headers.get('X-Page-To')
    if page_from and page_to:
        try:
            import fitz
            doc_pdf = fitz.open(input_path)
            total_pdf_pages = len(doc_pdf)
            
            p_from = int(page_from)
            p_to = int(page_to)
            
            p_from = max(1, min(p_from, total_pdf_pages))
            p_to = max(p_from, min(p_to, total_pdf_pages))
            
            sliced_doc = fitz.open()
            sliced_doc.insert_pdf(doc_pdf, from_page=p_from - 1, to_page=p_to - 1)
            doc_pdf.close()
            
            temp_sliced_path = input_path.replace('.pdf', '_sliced.pdf')
            sliced_doc.save(temp_sliced_path)
            sliced_doc.close()
            
            try:
                os.remove(input_path)
            except Exception:
                pass
            input_path = temp_sliced_path
        except Exception as se:
            logger.error(f"Slicing PDF range failed: {se}", exc_info=True)

    # ── Page Count Limit Check ──
    try:
        import fitz
        doc_pdf = fitz.open(input_path)
        total_pdf_pages = len(doc_pdf)
        doc_pdf.close()
    except Exception as e:
        logger.error(f"Failed to read PDF page count: {e}")
        try: os.remove(input_path)
        except Exception: pass
        return jsonify({'error': 'invalid_pdf', 'message': 'لا يمكن قراءة ملف PDF. يرجى التأكد من سلامة الملف.'}), 400

    limit_key = 'max_pages_premium' if effective_plan == 'premium' else 'max_pages_free'
    max_pages = int(get_setting(limit_key, '150' if effective_plan == 'premium' else '30'))

    if total_pdf_pages > max_pages:
        try: os.remove(input_path)
        except Exception: pass
        
        if effective_plan == 'premium':
            msg = f"يرجى تحديد نطاق الصفحات بحد أقصى 20 صفحة للحساب المميز (الملف المرفوع يحتوي على {total_pdf_pages} صفحات)."
        else:
            msg = f"يرجى تحديد نطاق الصفحات بحد أقصى 3 صفحات للحساب المجاني (الملف المرفوع يحتوي على {total_pdf_pages} صفحات)."
            
        return jsonify({
            'error': 'page_limit_exceeded',
            'message': msg,
            'limit': max_pages,
            'pages_found': total_pdf_pages
        }), 400

    output_filename = f"{conv_id}_{os.path.splitext(safe_name)[0]}.docx"
    output_path = os.path.join(OUTPUT_FOLDER, output_filename)

    # Get email and check concurrent conversions
    user_email = request.headers.get('X-User-Email')
    if not user_email and is_logged_in and user.email:
        user_email = user.email

    active_count = Conversion.query.filter_by(status='processing').count()
    limit_setting = AppSetting.query.filter_by(key='max_concurrent_conversions').first()
    max_concurrent = int(limit_setting.value if limit_setting else '10')

    cpu_percent, ram_percent = get_system_stats()
    # On Cloud Run, CPU percent might spike momentarily or report incorrectly, so rely only on max_concurrent
    should_queue = active_count >= max_concurrent

    conv_record = Conversion(
        conversion_id=conv_id,
        user_id=user.id,
        device_id=device_id,
        filename=safe_name,
        original_size=os.path.getsize(input_path),
        status='queued' if should_queue else 'processing',
        error_message='' if should_queue else '1/...',
        email=user_email if should_queue else None  # Only save email automatically if we are queueing
    )

    if should_queue:
        if os.getenv('GCS_BUCKET_NAME'):
            import gcs_helper
            gcs_input_name = f"inputs/{conv_id}_{safe_name}"
            gcs_helper.upload_to_gcs(input_path, gcs_input_name)
            
        # Save parameters for future execution by scheduler
        params = {
            'input_path': input_path,
            'gcs_input_name': f"inputs/{conv_id}_{safe_name}" if os.getenv('GCS_BUCKET_NAME') else None,
            'output_path': output_path,
            'language': language,
            'gemini_api_key': gemini_api_key,
            'conversion_mode': conversion_mode,
            'was_queued': True
        }
        conv_record.params = json.dumps(params)
        db.session.add(conv_record)
        db.session.commit()
        
        try:
            os.remove(input_path) # Free local space
        except Exception:
            pass

        # Calculate queue position
        position = Conversion.query.filter_by(status='queued').filter(Conversion.id < conv_record.id).count() + 1
        return jsonify({
            'success': True,
            'conversion_id': conv_id,
            'status': 'queued',
            'queue_position': position,
            'message': 'السيرفر مشغول حالياً. تم وضع ملفك في طابور الانتظار وسنعالجه قريباً.'
        })

    db.session.add(conv_record)
    db.session.commit()

    # Start conversion in a background thread
    import threading
    thread = threading.Thread(
        target=run_async_conversion,
        args=(conv_id, input_path, output_path, language, gemini_api_key, user.id, None, conversion_mode, None)
    )
    thread.daemon = True
    thread.start()

    return jsonify({
        'success': True,
        'conversion_id': conv_id,
        'status': 'processing'
    })


@app.route('/api/download/<conv_id>', methods=['GET'])
def download_file(conv_id):
    # Validate conv_id format to prevent path traversal
    if not validate_uuid(conv_id):
        return jsonify({'error': 'Invalid conversion ID'}), 400
    conv = Conversion.query.filter_by(conversion_id=conv_id).first()
    if not conv or conv.status != 'done':
        return jsonify({'error': 'File not found or not ready'}), 404

    if os.getenv('GCS_BUCKET_NAME'):
        import requests as req_lib
        from flask import Response
        bucket_name = os.getenv('GCS_BUCKET_NAME')
        
        # Build candidate filenames: try with original filename and generic names
        orig_name = conv.filename or 'document.pdf'
        base_name = os.path.splitext(orig_name)[0]
        candidates = [
            (f"outputs/{conv_id}_{base_name}.docx", f"{base_name}.docx", 'application/vnd.openxmlformats-officedocument.wordprocessingml.document'),
            (f"outputs/{conv_id}.docx", f"{base_name}.docx", 'application/vnd.openxmlformats-officedocument.wordprocessingml.document'),
            (f"outputs/{conv_id}_{base_name}.zip", f"{base_name}.zip", 'application/zip'),
            (f"outputs/{conv_id}.zip", f"{base_name}.zip", 'application/zip'),
        ]
        
        for gcs_path, dl_name, mime in candidates:
            url = f"https://storage.googleapis.com/{bucket_name}/{gcs_path}"
            r = req_lib.get(url, stream=True, timeout=10)
            if r.status_code == 200:
                logger.info(f"Serving download from GCS: {gcs_path}")
                return Response(
                    r.iter_content(chunk_size=8192),
                    content_type=mime,
                    headers={'Content-Disposition': f'attachment; filename="{dl_name}"'}
                )
        
        logger.warning(f"File not found in GCS for conv_id={conv_id}, tried all candidates")
        return jsonify({'error': 'File not found on storage'}), 404

    # Find the file locally
    for fname in os.listdir(OUTPUT_FOLDER):
        if fname.startswith(conv_id):
            file_path = os.path.join(OUTPUT_FOLDER, fname)
            mimetype = 'application/zip' if fname.endswith('.zip') else 'application/vnd.openxmlformats-officedocument.wordprocessingml.document'
            return send_file(
                file_path,
                as_attachment=True,
                download_name=fname.replace(f'{conv_id}_', ''),
                mimetype=mimetype
            )
    return jsonify({'error': 'File not found'}), 404


@app.route('/api/status/<conv_id>', methods=['GET'])
def conversion_status(conv_id):
    # Validate conv_id format to prevent injection/IDOR
    if not validate_uuid(conv_id):
        return jsonify({'error': 'Invalid conversion ID'}), 400
    conv = Conversion.query.filter_by(conversion_id=conv_id).first()
    if not conv:
        return jsonify({'error': 'Conversion not found'}), 404
        
    res = {
        'success': conv.status == 'done',
        'conversion_id': conv_id,
        'status': conv.status,
        'pages': conv.pages,
        'processing_time': conv.processing_time,
        'created_at': conv.created_at.isoformat(),
        'progress': conv.error_message if conv.status in ('processing', 'queued') else None,
        'engine': conv.error_message if conv.status == 'done' else None
    }
    if conv.status == 'done':
        res['download_url'] = f'/api/download/{conv_id}'
        res['filename'] = conv.filename.replace('.pdf', '.docx')
        for fname in os.listdir(OUTPUT_FOLDER):
            if fname.startswith(conv_id):
                res['filename'] = fname.replace(f'{conv_id}_', '')
                break
    elif conv.status == 'queued':
        position = Conversion.query.filter_by(status='queued').filter(Conversion.id < conv.id).count() + 1
        res['queue_position'] = position
    elif conv.status == 'failed':
        res['error'] = conv.error_message
        
    return jsonify(res)



@app.route('/api/queue/email', methods=['POST'])
def save_queue_email():
    try:
        data = request.get_json() or {}
        conv_id = data.get('conversion_id', '')
        email = data.get('email', '')
        
        if not conv_id or not email:
            return jsonify({'error': 'Missing parameters'}), 400
        
        # Validate formats
        if not validate_uuid(conv_id):
            return jsonify({'error': 'Invalid parameters'}), 400
        if not validate_email(email):
            return jsonify({'error': 'Invalid email format'}), 400
            
        conv = Conversion.query.filter_by(conversion_id=conv_id).first()
        if not conv:
            return jsonify({'error': 'Conversion not found'}), 404
            
        conv.email = email
        db.session.commit()
        
        # If the conversion was already finished while user was typing, send email immediately!
        if conv.status == 'done':
            site_url = os.getenv('SITE_URL', 'https://backend.hibdf.com').rstrip('/')
            download_url = f"{site_url}/api/download/{conv_id}"
            email_thread = threading.Thread(
                target=send_conversion_email,
                args=(conv.email, download_url, conv.filename)
            )
            email_thread.start()
            
        return jsonify({'success': True}), 200
    except Exception as e:
        logger.error(f"Failed to save queue email: {e}")
        return jsonify({'error': str(e)}), 500

@app.route('/api/user/activate-premium', methods=['POST'])
@jwt_required(optional=True)
def activate_premium():
    try:
        user_id = get_jwt_identity()
        data = request.get_json() or {}
        
        user = None
        if user_id:
            user = AppUser.query.get(user_id)
        
        if not user:
            # Fallback: find by device_id
            device_id = request.headers.get('X-Device-ID')
            if device_id:
                user = AppUser.query.filter_by(device_id=device_id).first()
        
        if not user:
            return jsonify({'error': 'User not found'}), 404
        
        # SECURITY: This endpoint should only be called from trusted backend processes
        # Verify a special server-side activation token
        activation_secret = os.getenv('ACTIVATION_SECRET', '')
        request_secret = request.headers.get('X-Activation-Secret', '')
        if not activation_secret or not hmac.compare_digest(activation_secret, request_secret):
            logger.warning(f"Unauthorized attempt to activate premium for user via frontend without secret")
            return jsonify({'error': 'Unauthorized activation'}), 403
        
        user.plan = 'premium'
        user.subscription_expiry = datetime.utcnow() + timedelta(days=31)
        db.session.commit()
        logger.info(f"User {user.id} activated premium via authorized callback")
        return jsonify({'success': True, 'plan': 'premium'}), 200
    except Exception as e:
        logger.error(f"activate_premium error: {e}")
        return jsonify({'error': str(e)}), 500

@app.route('/api/user/stats', methods=['GET'])
def user_stats():
    is_logged_in = False
    user = None
    device_id = request.headers.get('X-Device-ID', 'unknown')
    fingerprint = request.headers.get('X-Fingerprint')
    ip_address = request.headers.get('CF-Connecting-IP', request.headers.get('X-Forwarded-For', request.remote_addr))

    try:
        verify_jwt_in_request(optional=True)
        user_id = get_jwt_identity()
        if user_id:
            user = AppUser.query.get(user_id)
            if user:
                is_logged_in = True
    except:
        pass

    if not user:
        user = get_or_create_user(device_id, fingerprint, ip_address, 'ar', request.headers.get('CF-IPCountry'))

    free_limit = int(get_setting('free_lifetime_limit', '2'))
    if is_logged_in:
        free_limit += 2
        
    try:
        bonus = user.bonus_conversions if getattr(user, 'bonus_conversions', None) is not None else 0
        if is_logged_in:
            free_limit += bonus
    except:
        pass

    effective_plan = user.plan if is_logged_in else 'free'

    return jsonify({
        'id': user.id,
        'plan': effective_plan,
        'email': user.email or '',
        'lifetime_limit': free_limit if effective_plan == 'free' else -1,
        'total_conversions': user.total_conversions,
        'remaining_lifetime': max(0, free_limit - user.total_conversions) if effective_plan == 'free' else -1
    })

@app.route('/api/user/settings', methods=['GET'])
@jwt_required()
def user_settings():
    user_id = get_jwt_identity()
    user = AppUser.query.get(user_id)
    if not user:
        return jsonify({'error': 'User not found'}), 404
        
    return jsonify({
        'email': user.email or '',
        'plan': user.plan,
        'subscription_expiry': user.subscription_expiry.isoformat() if user.subscription_expiry else None,
        'auto_renew': user.auto_renew if hasattr(user, 'auto_renew') else True,
        'paddle_cancel_url': user.paddle_cancel_url if hasattr(user, 'paddle_cancel_url') else None
    })

@app.route('/api/user/fcm-token', methods=['POST'])
def save_fcm_token():
    user = get_user_from_request()
    if not user:
        return jsonify({'error': 'Unauthorized'}), 401
    
    data = request.json
    fcm_token = data.get('fcm_token')
    
    if fcm_token:
        user.fcm_token = fcm_token
        db.session.commit()
        return jsonify({'status': 'ok'})
    return jsonify({'error': 'Missing token'}), 400

@app.route('/api/user/reward-ad', methods=['POST'])
def reward_ad():
    try:
        verify_jwt_in_request(optional=True)
        user_id = get_jwt_identity()
        if user_id:
            user = AppUser.query.get(user_id)
        else:
            device_id = request.headers.get('X-Device-ID', 'unknown')
            fingerprint = request.headers.get('X-Fingerprint')
            ip_address = request.headers.get('CF-Connecting-IP', request.headers.get('X-Forwarded-For', request.remote_addr))
            user = get_or_create_user(device_id, fingerprint, ip_address, 'ar', request.headers.get('CF-IPCountry'))
    except:
        device_id = request.headers.get('X-Device-ID', 'unknown')
        fingerprint = request.headers.get('X-Fingerprint')
        ip_address = request.headers.get('CF-Connecting-IP', request.headers.get('X-Forwarded-For', request.remote_addr))
        user = get_or_create_user(device_id, fingerprint, ip_address, 'ar', request.headers.get('CF-IPCountry'))
        
    try:
        if getattr(user, 'bonus_conversions', None) is None:
            user.bonus_conversions = 0
        user.bonus_conversions += 2
        db.session.commit()
        return jsonify({'success': True, 'bonus_conversions': user.bonus_conversions})
    except Exception as e:
        return jsonify({'error': str(e)}), 500


@app.route('/api/visit', methods=['POST'])
def record_visit():
    data = request.get_json() or {}
    device_id = request.headers.get('X-Device-ID', 'unknown')
    page_name = sanitize_input(data.get('page', 'home'), 100)
    duration = data.get('duration', 0)
    try:
        duration = max(0, min(int(duration), 86400))  # Cap at 24 hours
    except (ValueError, TypeError):
        duration = 0
    visit = PageVisit(
        device_id=sanitize_input(device_id, 100),
        ip_address=request.headers.get('CF-Connecting-IP', request.headers.get('X-Forwarded-For', request.remote_addr)),
        page=page_name,
        duration_seconds=duration,
        country=sanitize_input(request.headers.get('CF-IPCountry', data.get('country', '')), 10)
    )
    db.session.add(visit)
    db.session.commit()
    return jsonify({'ok': True})


@app.route('/api/ads/config', methods=['GET'])
def get_ads_config():
    ads = Advertisement.query.filter_by(is_active=True).all()
    config = {}
    for ad in ads:
        config[ad.ad_type] = {
            'unit_id': ad.ad_unit_id,
            'enabled': ad.is_active,
            'position': ad.position or 'top',
            'video_url': ad.video_url or '',
            'ad_link': ad.ad_link or '',
            'ad_code': ad.ad_code or ''
        }
    return jsonify(config)


# ─────────────────────────────────────────────
# Admin Dashboard Routes
# ─────────────────────────────────────────────

@app.route('/api/admin/stats/overview', methods=['GET'])
@admin_required
def admin_stats_overview():
    today = datetime.utcnow().date()
    week_ago = datetime.utcnow() - timedelta(days=7)
    month_ago = datetime.utcnow() - timedelta(days=30)

    total_users = AppUser.query.count()
    active_today = AppUser.query.filter(
        AppUser.last_seen >= datetime.combine(today, datetime.min.time())
    ).count()
    total_conversions = Conversion.query.count()
    conversions_today = Conversion.query.filter(
        Conversion.created_at >= datetime.combine(today, datetime.min.time())
    ).count()
    premium_users = AppUser.query.filter_by(plan='premium').count()
    failed_conversions = Conversion.query.filter_by(status='failed').count()
    total_revenue = db.session.query(db.func.sum(Advertisement.revenue)).scalar() or 0

    # Daily conversions for last 7 days
    daily_data = []
    for i in range(7):
        day = datetime.utcnow().date() - timedelta(days=6-i)
        count = Conversion.query.filter(
            Conversion.created_at >= datetime.combine(day, datetime.min.time()),
            Conversion.created_at < datetime.combine(day + timedelta(days=1), datetime.min.time())
        ).count()
        daily_data.append({'date': day.isoformat(), 'conversions': count})

    cpu_percent, ram_percent = get_system_stats()
    active_conversions = Conversion.query.filter_by(status='processing').count()
    queued_conversions = Conversion.query.filter_by(status='queued').count()

    return jsonify({
        'total_users': total_users,
        'active_today': active_today,
        'total_conversions': total_conversions,
        'conversions_today': conversions_today,
        'premium_users': premium_users,
        'failed_conversions': failed_conversions,
        'total_revenue': round(total_revenue, 2),
        'free_users': total_users - premium_users,
        'daily_chart': daily_data,
        'cpu_usage': cpu_percent,
        'ram_usage': ram_percent,
        'active_conversions': active_conversions,
        'queued_conversions': queued_conversions
    })


@app.route('/api/admin/users', methods=['GET'])
@admin_required
def admin_get_users():
    page = request.args.get('page', 1, type=int)
    per_page = request.args.get('per_page', 20, type=int)
    search = request.args.get('search', '')
    plan_filter = request.args.get('plan', '')
    reg_filter = request.args.get('registered', '')

    query = AppUser.query
    if search:
        query = query.filter(AppUser.device_id.contains(search))
    if plan_filter:
        query = query.filter_by(plan=plan_filter)
    if reg_filter == 'registered':
        query = query.filter(AppUser.email.isnot(None))
    elif reg_filter == 'guest':
        query = query.filter(AppUser.email.is_(None))

    paginated = query.order_by(AppUser.created_at.desc()).paginate(page=page, per_page=per_page)

    return jsonify({
        'users': [{
            'id': u.id,
            'email': u.email,
            'device_id': (u.device_id[:20] + '...') if u.device_id else None,
            'platform': u.platform,
            'language': u.language,
            'plan': u.plan,
            'total_conversions': u.total_conversions,
            'created_at': u.created_at.isoformat(),
            'last_seen': u.last_seen.isoformat() if u.last_seen else None,
            'country': u.country,
            'ip_address': u.ip_address or '—'
        } for u in paginated.items],
        'total': paginated.total,
        'pages': paginated.pages,
        'current_page': page
    })


@app.route('/api/admin/users/<int:user_id>/upgrade', methods=['POST'])
@admin_required
def admin_upgrade_user(user_id):
    user = AppUser.query.get_or_404(user_id)
    data = request.get_json()
    new_plan = data.get('plan', 'premium')
    if new_plan not in ('free', 'premium'):
        return jsonify({'error': 'Invalid plan value'}), 400
    user.plan = new_plan
    db.session.commit()
    return jsonify({'success': True, 'plan': user.plan})


@app.route('/api/admin/conversions', methods=['GET'])
@admin_required
def admin_get_conversions():
    page = request.args.get('page', 1, type=int)
    per_page = request.args.get('per_page', 20, type=int)
    status_filter = request.args.get('status', '')

    query = Conversion.query
    if status_filter:
        query = query.filter_by(status=status_filter)

    paginated = query.order_by(Conversion.created_at.desc()).paginate(page=page, per_page=per_page)

    return jsonify({
        'conversions': [{
            'id': c.id,
            'conversion_id': c.conversion_id[:16] + '...',
            'filename': c.filename,
            'pages': c.pages,
            'language': c.language_detected,
            'status': c.status,
            'processing_time': c.processing_time,
            'created_at': c.created_at.isoformat()
        } for c in paginated.items],
        'total': paginated.total,
        'pages': paginated.pages
    })


@app.route('/api/admin/ads', methods=['GET'])
@admin_required
def admin_get_ads():
    ads = Advertisement.query.all()
    return jsonify([{
        'id': a.id,
        'name': a.name,
        'ad_type': a.ad_type,
        'ad_unit_id': a.ad_unit_id,
        'position': a.position or 'top',
        'video_url': a.video_url or '',
        'ad_link': a.ad_link or '',
        'is_active': a.is_active,
        'impressions': a.impressions,
        'clicks': a.clicks,
        'revenue': a.revenue,
        'ctr': round((a.clicks / a.impressions * 100) if a.impressions > 0 else 0, 2),
        'created_at': a.created_at.isoformat()
    } for a in ads])


@app.route('/api/admin/ads', methods=['POST'])
@admin_required
def admin_create_ad():
    data = request.get_json()
    ad = Advertisement(
        name=data.get('name'),
        ad_type=data.get('ad_type'),
        ad_unit_id=data.get('ad_unit_id', ''),
        position=data.get('position', 'top'),
        video_url=data.get('video_url', ''),
        ad_link=data.get('ad_link', ''),
        ad_code=data.get('ad_code', ''),
        is_active=data.get('is_active', True)
    )
    db.session.add(ad)
    db.session.commit()
    return jsonify({'success': True, 'id': ad.id}), 201


@app.route('/api/admin/ads/<int:ad_id>', methods=['PUT'])
@admin_required
def admin_update_ad(ad_id):
    ad = Advertisement.query.get_or_404(ad_id)
    data = request.get_json()
    for field in ['name', 'ad_type', 'ad_unit_id', 'position', 'video_url', 'ad_link', 'ad_code', 'is_active', 'revenue']:
        if field in data:
            setattr(ad, field, data[field])
    ad.updated_at = datetime.utcnow()
    db.session.commit()
    return jsonify({'success': True})


@app.route('/api/admin/ads/<int:ad_id>', methods=['DELETE'])
@admin_required
def admin_delete_ad(ad_id):
    ad = Advertisement.query.get_or_404(ad_id)
    db.session.delete(ad)
    db.session.commit()
    return jsonify({'success': True})


@app.route('/api/admin/upload-ad-media', methods=['POST'])
@admin_required
def admin_upload_ad_media():
    if 'file' not in request.files:
        return jsonify({'error': 'No file part'}), 400
    file = request.files['file']
    if file.filename == '':
        return jsonify({'error': 'No selected file'}), 400
        
    ext = os.path.splitext(file.filename)[1].lower()
    if ext not in ['.png', '.jpg', '.jpeg', '.gif', '.mp4', '.webm']:
        return jsonify({'error': 'Invalid file type'}), 400
        
    import time
    filename = f"ad_{int(time.time())}{ext}"
    os.makedirs(os.path.join(app.root_path, 'static', 'ads'), exist_ok=True)
    
    # Save the file locally using os.path.join for Windows compatibility
    file_path = os.path.join(app.root_path, 'static', 'ads', filename)
    file.save(file_path)
    
    bucket_name = os.getenv('GCS_BUCKET_NAME')
    if bucket_name:
        try:
            from gcs_helper import get_gcs_client
            client = get_gcs_client()
            if client:
                bucket = client.bucket(bucket_name)
                blob = bucket.blob(f"ads/{filename}")
                blob.upload_from_filename(file_path)
                # Always use forward slashes for URLs
                return jsonify({'url': f"https://storage.googleapis.com/{bucket_name}/ads/{filename}"}), 200
        except Exception as e:
            app.logger.error(f"GCS upload failed: {e}")
            
    # Fallback to local path URL using forward slash
    return jsonify({'url': f"/static/ads/{filename}"}), 200


@app.route('/api/admin/settings', methods=['GET'])
@admin_required
def admin_get_settings():
    settings = AppSetting.query.all()
    return jsonify({s.key: s.value for s in settings})


@app.route('/api/admin/fix-stuck', methods=['GET'])
@admin_required
def admin_fix_stuck():
    try:
        updated = Conversion.query.filter(Conversion.status.in_(['processing', 'queued'])).update({
            'status': 'failed',
            'error_message': 'Cleared by admin'
        }, synchronize_session=False)
        db.session.commit()
        return jsonify({'success': True, 'cleared_count': updated})
    except Exception as e:
        return jsonify({'error': str(e)}), 500

@app.route('/api/admin/migrate', methods=['GET'])
@admin_required
def admin_migrate_db():
    from sqlalchemy import text
    results = []
    try:
        db.session.execute(text("ALTER TABLE advertisements ADD COLUMN position VARCHAR(100) DEFAULT 'top'"))
        results.append("Added 'position' column.")
    except Exception as e:
        # SQLite/Postgres will throw an error if column already exists
        results.append("position column already exists or error.")
        
    try:
        db.session.execute(text("ALTER TABLE advertisements ADD COLUMN video_url VARCHAR(500)"))
        results.append("Added 'video_url' column.")
    except Exception as e:
        results.append("video_url column already exists or error.")
        
    try:
        db.session.execute(text("ALTER TABLE advertisements ADD COLUMN ad_link VARCHAR(500)"))
        results.append("Added 'ad_link' column.")
    except Exception as e:
        results.append("ad_link column already exists or error.")
        
    try:
        db.session.execute(text("ALTER TABLE advertisements ADD COLUMN ad_code TEXT"))
        results.append("Added 'ad_code' column.")
    except Exception as e:
        results.append("ad_code column already exists or error.")
        
    db.session.commit()
    return jsonify({"success": True, "details": results})

@app.route('/api/admin/settings', methods=['POST'])
@admin_required
def admin_update_settings():
    data = request.get_json()
    for key, value in data.items():
        setting = AppSetting.query.filter_by(key=key).first()
        if setting:
            setting.value = str(value)
        else:
            setting = AppSetting(key=key, value=str(value))
            db.session.add(setting)
    db.session.commit()
    return jsonify({'success': True})

@app.route('/api/admin/change-password', methods=['POST'])
@admin_required
def admin_change_password():
    data = request.get_json()
    current_password = data.get('current_password')
    new_password = data.get('new_password')
    
    if not current_password or not new_password:
        return jsonify({'error': 'يرجى إدخال كلمة المرور الحالية والجديدة'}), 400
        
    username = get_jwt_identity()
    admin = AdminUser.query.filter_by(username=username).first()
    
    if not admin.check_password(current_password):
        return jsonify({'error': 'كلمة المرور الحالية غير صحيحة'}), 403
        
    admin.set_password(new_password)
    db.session.commit()
    
    return jsonify({'success': True, 'message': 'تم تغيير كلمة المرور بنجاح!'})


@app.route('/api/admin/analytics/languages', methods=['GET'])
@admin_required
def admin_language_stats():
    result = db.session.query(
        AppUser.language, db.func.count(AppUser.id)
    ).group_by(AppUser.language).all()
    return jsonify([{'language': r[0], 'count': r[1]} for r in result])


@app.route('/api/admin/analytics/countries', methods=['GET'])
@admin_required
def admin_country_stats():
    result = db.session.query(
        AppUser.country, db.func.count(AppUser.id)
    ).group_by(AppUser.country).filter(AppUser.country != None).all()
    return jsonify([{'country': r[0], 'count': r[1]} for r in result])

@app.route('/api/admin/analytics/realtime', methods=['GET'])
@admin_required
def admin_realtime_stats():
    five_mins_ago = datetime.utcnow() - timedelta(minutes=5)
    
    # Get recent app users
    app_users_ips = [u.ip_address for u in AppUser.query.filter(AppUser.last_seen >= five_mins_ago).all() if u.ip_address]
    
    # Get recent visitors
    recent_visits = PageVisit.query.filter(PageVisit.created_at >= five_mins_ago).all()
    visit_ips = [v.ip_address for v in recent_visits if getattr(v, 'ip_address', None)]
    
    # Count unique IPs
    unique_online_ips = len(set(app_users_ips + visit_ips))
    if unique_online_ips == 0:
        # Fallback to pure count if IP is empty
        unique_online_ips = AppUser.query.filter(AppUser.last_seen >= five_mins_ago).count() + PageVisit.query.filter(PageVisit.created_at >= five_mins_ago).count()
    
    active_conversions = Conversion.query.filter_by(status='processing').count()
    
    return jsonify({
        'online_users': unique_online_ips,
        'active_conversions': active_conversions
    })

@app.route('/api/admin/analytics/financials', methods=['GET'])
@admin_required
def admin_financial_stats():
    # Example logic: Summarize PaymentTransaction by month/week
    # For now, we will return dummy/placeholder data grouped appropriately
    # In a real scenario, we'd run a GROUP BY over PaymentTransaction.created_at
    
    # Yearly, Monthly, Weekly revenue
    yearly = db.session.query(db.func.sum(PaymentTransaction.amount)).scalar() or 0.0
    
    month_ago = datetime.utcnow() - timedelta(days=30)
    monthly = db.session.query(db.func.sum(PaymentTransaction.amount)).filter(PaymentTransaction.created_at >= month_ago).scalar() or 0.0
    
    week_ago = datetime.utcnow() - timedelta(days=7)
    weekly = db.session.query(db.func.sum(PaymentTransaction.amount)).filter(PaymentTransaction.created_at >= week_ago).scalar() or 0.0
    
    return jsonify({
        'yearly': yearly,
        'monthly': monthly,
        'weekly': weekly
    })

@app.cli.command("create-admin")
def create_admin():
    import click
    username = click.prompt("Enter admin username")
    email = click.prompt("Enter admin email")
    password = click.prompt("Enter admin password", hide_input=True)
    
    if AdminUser.query.filter_by(username=username).first():
        click.echo("Admin already exists!")
        return
        
    admin = AdminUser(username=username, email=email)
    admin.set_password(password)
    db.session.add(admin)
    db.session.commit()
    click.echo(f"Admin {username} created successfully!")


# ─────────────────────────────────────────────
# Init DB + Default Data
# ─────────────────────────────────────────────

# ─────────────────────────────────────────────
# Webhooks
# ─────────────────────────────────────────────

_last_client_alert = {'t': 0.0}

@app.route('/api/stats', methods=['GET'])
@app.route('/api/health', methods=['GET'])
def health_check():
    return jsonify({'status': 'ok', 'service': 'hibdf-backend'}), 200

@app.route('/api/client-alert', methods=['POST'])
def client_alert():
    """Relay frontend error alerts to Telegram. Token stays server-side (env TELEGRAM_BOT_TOKEN)."""
    import time as _time
    import urllib.request as _ur
    import urllib.parse as _up
    token = os.getenv('TELEGRAM_BOT_TOKEN', '')
    chat_id = os.getenv('TELEGRAM_CHAT_ID', '')
    if not token or not chat_id:
        return jsonify({'ok': False}), 200
    now = _time.time()
    if now - _last_client_alert['t'] < 30:  # simple global rate limit
        return jsonify({'ok': False, 'throttled': True}), 200
    _last_client_alert['t'] = now
    data = request.get_json(silent=True) or {}
    reason = str(data.get('reason', ''))[:300]
    msg = "\U0001F6A8 hibdf alert\n" + reason
    try:
        url = f"https://api.telegram.org/bot{token}/sendMessage?" + _up.urlencode({'chat_id': chat_id, 'text': msg})
        _ur.urlopen(url, timeout=5)
    except Exception as e:
        logger.warning(f"client_alert telegram failed: {e}")
    return jsonify({'ok': True}), 200

@app.route('/api/paddle/webhook', methods=['POST'])
def paddle_webhook():
    try:
        # Verify Paddle Signature for Security
        paddle_secret = os.getenv('PADDLE_WEBHOOK_SECRET', '')
        if paddle_secret:
            signature_header = request.headers.get('Paddle-Signature')
            if not signature_header:
                logger.warning("Paddle Webhook rejected: Missing Signature")
                return jsonify({'error': 'Unauthorized - Missing Signature'}), 401
                
            parts = dict(x.split('=', 1) for x in signature_header.split(';') if '=' in x)
            ts = parts.get('ts', '')
            h1 = parts.get('h1', '')
            
            # Validate timestamp to prevent replay attacks (5 minute window)
            try:
                webhook_ts = int(ts)
                current_ts = int(datetime.utcnow().timestamp())
                if abs(current_ts - webhook_ts) > 300:  # 5 minutes
                    logger.warning("Paddle Webhook rejected: Timestamp too old (replay attack prevention)")
                    return jsonify({'error': 'Unauthorized - Timestamp expired'}), 401
            except (ValueError, TypeError):
                return jsonify({'error': 'Unauthorized - Invalid timestamp'}), 401
            
            raw_body = request.get_data(as_text=True)
            signed_payload = ts + ':' + raw_body
            mac = hmac.new(paddle_secret.encode('utf-8'), signed_payload.encode('utf-8'), hashlib.sha256).hexdigest()
            
            if not hmac.compare_digest(mac, h1):
                logger.warning("Paddle Webhook rejected: Invalid Signature")
                return jsonify({'error': 'Unauthorized - Invalid Signature'}), 401

        payload = request.json
        if not payload:
            return jsonify({'error': 'No payload'}), 400
            
        event_type = payload.get('event_type')
        data = payload.get('data', {})
        
        logger.info(f"Received Paddle Webhook: {event_type}")
        
        if event_type in ['subscription.created', 'subscription.updated']:
            status = data.get('status')
            custom_data = data.get('custom_data', {})
            user_id = custom_data.get('user_id')
            
            if user_id and status == 'active':
                user = AppUser.query.get(user_id)
                if user:
                    # Determine subscription duration from billing cycle
                    items = data.get('items', [])
                    days_to_add = 31  # default monthly
                    if items:
                        interval = items[0].get('price', {}).get('billing_cycle', {}).get('interval', 'month')
                        frequency = items[0].get('price', {}).get('billing_cycle', {}).get('frequency', 1)
                        if interval == 'year':
                            days_to_add = 366 * frequency
                        else:
                            days_to_add = 31 * frequency
                    
                    # Stack time: if already premium with future expiry, add on top
                    now = datetime.utcnow()
                    if user.plan == 'premium' and user.subscription_expiry and user.subscription_expiry > now:
                        # Add new period on top of remaining time
                        user.subscription_expiry = user.subscription_expiry + timedelta(days=days_to_add)
                        logger.info(f"User {user_id} renewed - stacked +{days_to_add} days, new expiry: {user.subscription_expiry}")
                    else:
                        # New subscription or expired - start fresh
                        user.subscription_expiry = now + timedelta(days=days_to_add)
                        logger.info(f"User {user_id} new subscription - {days_to_add} days, expiry: {user.subscription_expiry}")
                    
                    user.plan = 'premium'
                    user.paddle_subscription_id = data.get('id')
                    cancel_url = data.get('management_urls', {}).get('cancel')
                    if cancel_url:
                        user.paddle_cancel_url = cancel_url
                    
                    db.session.commit()
                    
        elif event_type in ['subscription.canceled', 'subscription.past_due']:
            custom_data = data.get('custom_data', {})
            user_id = custom_data.get('user_id')
            if user_id:
                user = AppUser.query.get(user_id)
                if user:
                    # On cancel: keep premium until expiry date (don't cut off immediately)
                    if event_type == 'subscription.canceled':
                        user.auto_renew = False
                        # Let them keep premium until subscription_expiry naturally
                        logger.info(f"User {user_id} canceled - will remain premium until {user.subscription_expiry}")
                    else:
                        # Past due: downgrade immediately
                        user.plan = 'free'
                        logger.info(f"User {user_id} subscription past_due - downgraded to free")
                    db.session.commit()
                    
        return jsonify({'status': 'ok'}), 200
    except Exception as e:
        logger.error(f"Paddle webhook error: {e}", exc_info=True)
        return jsonify({'error': 'Webhook processing error'}), 500

def send_support_email(user_email, user_message, device_id):
    import smtplib
    from email.mime.text import MIMEText
    from email.mime.multipart import MIMEMultipart

    smtp_server = os.getenv('SMTP_SERVER', '').strip()
    smtp_port = os.getenv('SMTP_PORT', '587').strip()
    smtp_user = os.getenv('SMTP_USER', '').strip()
    smtp_pass = os.getenv('SMTP_PASSWORD', '').strip()
    smtp_from = os.getenv('SMTP_FROM_EMAIL', 'PDF Pro <noreply@hibdf.com>').strip()
    admin_email = 'moh.salameh2992@gmail.com'
    if not smtp_server or not smtp_user or not smtp_pass:
        logger.warning(f"SMTP not configured. Skipping support email from {user_email}")
        return False, "SMTP variables missing in .env"

    try:
        msg = MIMEMultipart('alternative')
        msg['Subject'] = f"Support Request from {user_email}"
        msg['From'] = smtp_from
        msg['To'] = admin_email
        msg.add_header('reply-to', user_email)

        html = f"""
        <html>
        <body style="font-family: Arial, sans-serif; color: #333; line-height: 1.6;">
            <h2 style="color: #a855f7;">New Support Request</h2>
            <p><strong>From:</strong> {user_email}</p>
            <p><strong>Device ID:</strong> {device_id}</p>
            <hr>
            <p><strong>Message:</strong></p>
            <p style="white-space: pre-wrap; background: #f4f4f4; padding: 15px; border-radius: 5px;">{user_message}</p>
        </body>
        </html>
        """
        part = MIMEText(html, 'html')
        msg.attach(part)

        server = smtplib.SMTP(smtp_server, int(smtp_port))
        server.starttls()
        server.login(smtp_user, smtp_pass)
        server.sendmail(smtp_from, [admin_email], msg.as_string())
        server.quit()
        logger.info(f"Successfully sent support email from {user_email}")
        return True, ""
    except Exception as e:
        logger.error(f"Failed to send support email from {user_email}: {e}")
        return False, str(e)

@app.route('/api/contact', methods=['GET', 'POST'])
@app.route('/api/api/contact', methods=['GET', 'POST'])
@app.route('/api/user/support', methods=['POST'])
def contact_support():
    if request.method == 'GET':
        return jsonify({'error': 'GET request received but POST expected. The proxy converted the POST to GET!'}), 500
    
    data = request.get_json()
    email = data.get('email')
    message = data.get('message')
    device_id = request.headers.get('X-Device-ID', 'unknown')

    if not email or not message:
        return jsonify({'error': 'Email and message are required'}), 400
    
    # Validate email format
    if not validate_email(email):
        return jsonify({'error': 'Invalid email format'}), 400
    
    # Limit message length
    message = str(message)[:5000]

    try:
        success, err_msg = send_support_email(email, message, device_id)
        if success:
            return jsonify({'success': True}), 200
        else:
            logger.error(f"Support email failed: {err_msg}")
            return jsonify({'error': 'Failed to send message. Please try again.'}), 500
    except Exception as e:
        logger.error(f"Support route exception: {e}")
        return jsonify({'error': 'An error occurred. Please try again.'}), 500

@app.errorhandler(400)
def bad_request(e):
    return jsonify({'error': 'Bad request'}), 400

@app.errorhandler(404)
def not_found(e):
    return jsonify({'error': 'Not found'}), 404

@app.errorhandler(405)
def method_not_allowed(e):
    return jsonify({'error': 'Method not allowed'}), 405

@app.errorhandler(500)
def internal_error(e):
    logger.error(f"Internal server error: {e}", exc_info=True)
    return jsonify({'error': 'Internal server error'}), 500

@app.errorhandler(429)
def rate_limit_exceeded(e):
    return jsonify({'error': 'Too many requests. Please slow down.', 'retry_after': 60}), 429

def init_db():
    with app.app_context():
        db.create_all()

        # Create default admin
        try:
            if not AdminUser.query.filter_by(username='admin').first():
                admin = AdminUser(
                    username='admin',
                    email='admin@pdfpro.app',
                    role='superadmin'
                )
                admin.set_password('Admin@2026!')
                db.session.add(admin)
                db.session.commit()
                logger.info("✅ Default admin created: admin / Admin@2026!")
        except Exception:
            db.session.rollback()

        # Default settings
        defaults = {
            'free_daily_limit': '5',
            'app_name': 'PDF Pro Converter',
            'app_version': '1.0.0',
            'maintenance_mode': 'false',
            'premium_price_usd': '4.99',
            'max_pages_free': '3',
            'max_pages_premium': '20',
            'max_concurrent_conversions': '10'
        }
        for key, value in defaults.items():
            try:
                setting = AppSetting.query.filter_by(key=key).first()
                if setting:
                    # Force update limits to apply the new strict page limits
                    if key in ('max_pages_free', 'max_pages_premium'):
                        setting.value = value
                        db.session.commit()
                else:
                    db.session.add(AppSetting(key=key, value=value))
                    db.session.commit()
            except Exception:
                db.session.rollback()

        # Default ads
        try:
            if Advertisement.query.count() == 0:
                ads = [
                    Advertisement(name='Banner Ad - Main', ad_type='banner', ad_unit_id='ca-app-pub-TEST/BANNER001', is_active=True),
                    Advertisement(name='Interstitial - After Conversion', ad_type='interstitial', ad_unit_id='ca-app-pub-TEST/INTERSTITIAL001', is_active=True),
                    Advertisement(name='Rewarded - Extra Conversions', ad_type='rewarded', ad_unit_id='ca-app-pub-TEST/REWARDED001', is_active=True),
                ]
                for ad in ads:
                    db.session.add(ad)
                db.session.commit()
        except Exception:
            db.session.rollback()

        try:
            db.session.commit()
            logger.info("✅ Database initialized successfully")
        except Exception:
            db.session.rollback()
        
        global _schedulers_started
        if not _schedulers_started:
            start_queue_scheduler()
            start_file_cleanup_scheduler()
            _schedulers_started = True


_schedulers_started = False
init_db()

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=8000, debug=False)
