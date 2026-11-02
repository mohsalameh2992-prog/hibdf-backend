"""
PDF Pro Converter v5.0 — The "World-Class" Commercial Engine
============================================================
This engine uses a hybrid commercial-grade approach to guarantee 100% layout fidelity.

Pipeline:
1. Detect PDF Type (Digital vs. Scanned)
2. If Scanned: Rebuild the PDF using Tesseract to inject a perfectly aligned, 
   searchable hidden text layer directly into the PDF.
3. Convert to DOCX: Use Spire.Pdf (a commercial-grade C++ layout engine) to 
   convert the PDF directly into a Word Document. This ensures margins, 
   colors, images, and text positioning are mathematically identical to the original.
4. Clean: Use python-docx to remove evaluation watermarks automatically.
"""

import os, io, sys, logging, time
import fitz
from PIL import Image, ImageEnhance, ImageFilter
from docx import Document

# ─── TESSERACT ───
import pytesseract
for p in [
    r'/usr/bin/tesseract',                                    # Linux / Docker
    r'/usr/local/bin/tesseract',                             # macOS / some Linux
    r'C:\Program Files\Tesseract-OCR\tesseract.exe',         # Windows
    r'C:\Program Files (x86)\Tesseract-OCR\tesseract.exe',   # Windows x86
]:
    if os.path.exists(p):
        pytesseract.pytesseract.tesseract_cmd = p
        break

# ─── SPIRE.PDF ───
from spire.pdf.common import *
from spire.pdf import *

logger = logging.getLogger(__name__)

import itertools
import threading
import re

_api_key_cycles = {}
_api_key_lock = threading.Lock()

def get_next_api_key(api_key_string):
    global _api_key_cycles, _api_key_lock
    if not api_key_string: return ""
    keys = [k.strip() for k in re.split(r'[,;\s]+', api_key_string) if k.strip()]
    if not keys: return ""
    
    with _api_key_lock:
        # Use the raw api_key_string as the dictionary key to track independent cycles
        if api_key_string not in _api_key_cycles:
            _api_key_cycles[api_key_string] = itertools.cycle(keys)
        return next(_api_key_cycles[api_key_string])

class PDFConverter:
    def __init__(self, tesseract_path=None):
        if tesseract_path and os.path.exists(tesseract_path):
            pytesseract.pytesseract.tesseract_cmd = tesseract_path

    def _detect_is_arabic(self, pdf_doc, user_lang, file_path):
        """Returns True if the document should be treated as Arabic/RTL."""
        if user_lang == 'ar':
            return True
        if user_lang == 'en':
            return False
        # lang == 'auto' → sample text from first 3 pages
        arabic_chars = 0
        total_chars  = 0
        for idx in range(min(3, len(pdf_doc))):
            txt = pdf_doc[idx].get_text('text')
            for ch in txt:
                cp = ord(ch)
                if 0x0600 <= cp <= 0x06FF or 0xFB50 <= cp <= 0xFDFF or 0xFE70 <= cp <= 0xFEFF:
                    arabic_chars += 1
                if not ch.isspace():
                    total_chars += 1
        if total_chars > 20:
            is_ar = (arabic_chars / total_chars) >= 0.10
            logger.info(f"Language auto-detect: {total_chars} chars found. Arabic chars: {arabic_chars}. is_arabic={is_ar}")
            return is_ar
        # Scanned/image-only PDF → fall back to filename keywords
        name_lower = os.path.basename(file_path).lower()
        en_keywords = ['en', 'eng', 'english', 'activity', 'worksheet', 'exam']
        if any(kw in name_lower for kw in en_keywords):
            logger.info(f"Language auto-detect: Scanned PDF, English keyword found in filename '{name_lower}'. is_arabic=False")
            return False
        logger.info(f"Language auto-detect: Scanned PDF '{name_lower}', no English keywords. Defaulting is_arabic=True")
        return True

    def _pre_process_heavy_pages(self, input_pdf_path: str, output_pdf_path: str) -> bool:
        """Detects if there are heavy pages in the PDF (>40 images) and flattens them to high-resolution images to prevent Spire from hanging.
        Returns True if any pages were flattened, False otherwise."""
        doc_src = fitz.open(input_pdf_path)
        has_heavy_pages = False
        for page in doc_src:
            if len(page.get_images()) > 40:
                has_heavy_pages = True
                break
                
        if not has_heavy_pages:
            doc_src.close()
            return False
            
        logger.info(f"Flattening heavy pages in {input_pdf_path} to prevent Spire hanging...")
        processed_doc = fitz.open()
        for idx, page in enumerate(doc_src):
            num_imgs = len(page.get_images())
            if num_imgs > 40:
                logger.info(f"Page {idx+1} has {num_imgs} images. Rasterizing...")
                # Render page at 288 DPI (zoom=4) for crisp text quality
                pix = page.get_pixmap(matrix=fitz.Matrix(4, 4))
                img_data = pix.tobytes("png")
                
                # Create a 1-page PDF from the PNG data
                img_doc = fitz.open("png", img_data)
                pdf_bytes = img_doc.convert_to_pdf()
                img_doc.close()
                
                single_page_pdf = fitz.open("pdf", pdf_bytes)
                processed_doc.insert_pdf(single_page_pdf)
                single_page_pdf.close()
            else:
                processed_doc.insert_pdf(doc_src, from_page=idx, to_page=idx)
                
        processed_doc.save(output_pdf_path)
        processed_doc.close()
        doc_src.close()
        return True

    def convert(self, input_path: str, output_path: str,
                language: str = 'auto', mode: str = 'auto', gemini_api_key: str = '',
                progress_callback=None) -> dict:
        
        start_time = time.time()
        doc_pdf = fitz.open(input_path)
        total_pages = len(doc_pdf)
        is_scanned = self._is_scanned(doc_pdf)
        is_arabic = self._detect_is_arabic(doc_pdf, language, input_path)
        doc_pdf.close()

        # Engine decision logic
        use_ai = False
        gemini_error = None
        if gemini_api_key:
            if mode == 'force_local':
                use_ai = False
            elif mode == 'force_ai':
                use_ai = True
            else:
                # mode == 'auto'
                # Use Gemini ONLY for scanned/image Arabic PDFs.
                # Digital Arabic PDFs (with real text) are handled better by the local engine.
                use_ai = is_arabic and is_scanned

        logger.info(f"Engine selection -> is_scanned: {is_scanned}, is_arabic: {is_arabic}, mode: {mode}, has_key: {bool(gemini_api_key)}, use_ai: {use_ai}")

        # --- MULTI-PAGE SPLIT & MERGE LOGIC ---
        if total_pages > 1:
            logger.info(f"File has {total_pages} pages. Splitting and converting page-by-page, then merging...")
            import uuid as uuid_mod
            import subprocess, sys
            import cv2
            import numpy as np

            base_out, _ = os.path.splitext(output_path)
            doc_src = fitz.open(input_path)
            chunk_script = os.path.join(os.path.dirname(os.path.abspath(__file__)), "convert_chunk.py")

            from concurrent.futures import ThreadPoolExecutor
            import threading

            api_key_groups = [g.strip() for g in gemini_api_key.split('|') if g.strip()]
            
            completed_lock = threading.Lock()
            pdf_extract_lock = threading.Lock()
            completed_count = 0
            gemini_success_count = 0
            converted_results = {}

            def convert_page_worker(idx):
                nonlocal completed_count, gemini_success_count
                page_num = idx + 1
                uid = uuid_mod.uuid4().hex[:6]
                chunk_pdf = f"{base_out}_sub_{idx}_{uid}.pdf"
                chunk_docx = f"{base_out}_sub_{idx}_{uid}.docx"

                # Save single page PDF (thread-safe)
                with pdf_extract_lock:
                    chunk_doc = fitz.open()
                    chunk_doc.insert_pdf(doc_src, from_page=idx, to_page=idx)
                    chunk_doc.save(chunk_pdf)
                    chunk_doc.close()

                # Assign this page to a specific API key group (e.g. specific email project)
                group_string = api_key_groups[idx % len(api_key_groups)] if api_key_groups else gemini_api_key

                converted = False
                is_gemini_success = False

                # Convert via Gemini Vision (if enabled)
                if use_ai:
                    try:
                        logger.info(f"Converting page {page_num} via Gemini Vision...")
                        self._convert_via_gemini_vision(chunk_pdf, chunk_docx, language, group_string)
                        if os.path.exists(chunk_docx) and os.path.getsize(chunk_docx) > 500:
                            converted = True
                            is_gemini_success = True
                            logger.info(f"Page {page_num}: Gemini Vision succeeded.")
                        else:
                            logger.warning(f"Page {page_num}: Gemini Vision produced empty output. Falling back to local engine.")
                    except RuntimeError as re_err:
                        logger.error(f"Gemini Vision FATAL error for page {page_num}: {re_err}. Aborting conversion.")
                        raise re_err
                    except Exception as ge:
                        logger.warning(f"Gemini Vision failed for page {page_num}: {ge}. Falling back to local engine.")

                # Convert via local engine (if Gemini is disabled or failed)
                if not converted:
                    local_pdf = chunk_pdf
                    temp_chunk_ocr = chunk_pdf.replace('.pdf', '_ocr.pdf')
                    
                    page_is_scanned = is_scanned
                    if not is_scanned and mode != 'ocr':
                        try:
                            page_doc = fitz.open(chunk_pdf)
                            page_text = page_doc[0].get_text('text').strip() if len(page_doc) > 0 else ''
                            page_doc.close()
                            page_is_scanned = len(page_text) < 30
                        except Exception:
                            page_is_scanned = False
                    
                    if page_is_scanned or mode == 'ocr':
                        try:
                            logger.info(f"Page {page_num} is image-only. Running OCR to extract text...")
                            import pytesseract
                            from docx import Document as DocxDocument
                            
                            img_doc = fitz.open(local_pdf)
                            page = img_doc.load_page(0)
                            pix = page.get_pixmap(matrix=fitz.Matrix(2, 2))
                            img_data = np.frombuffer(pix.samples, dtype=np.uint8).reshape((pix.h, pix.w, 3))
                            img_gray = cv2.cvtColor(img_data, cv2.COLOR_RGB2GRAY)
                            img_doc.close()
                            
                            tess_lang = 'ara+eng' if is_arabic else 'eng'
                            ocr_text = pytesseract.image_to_string(img_gray, lang=tess_lang, config='--oem 3 --psm 4').strip()
                            
                            word_doc = DocxDocument()
                            is_rtl_page = language in ('auto', 'ar', '')
                            if ocr_text:
                                from docx.enum.text import WD_ALIGN_PARAGRAPH
                                for line in ocr_text.split('\n'):
                                    para = word_doc.add_paragraph(line)
                                    para.alignment = WD_ALIGN_PARAGRAPH.RIGHT if is_rtl_page else WD_ALIGN_PARAGRAPH.LEFT
                            word_doc.save(chunk_docx)
                            
                            if os.path.exists(chunk_docx) and os.path.getsize(chunk_docx) > 0:
                                converted = True
                                logger.info(f"Page {page_num}: Local OCR conversion succeeded.")
                        except Exception as oe:
                            logger.warning(f"Local OCR failed for page {page_num}: {oe}")
                    
                    if not converted:
                        try:
                            logger.info(f"Running subprocess fallback for page {page_num}...")
                            res = subprocess.run(
                                [sys.executable, chunk_script, local_pdf, chunk_docx, 'ar' if is_arabic else 'en'],
                                capture_output=True, text=True, timeout=120
                            )
                            if res.returncode == 0 and os.path.exists(chunk_docx) and os.path.getsize(chunk_docx) > 0:
                                converted = True
                            else:
                                logger.warning(f"Subprocess failed for page {page_num}: {res.stderr[:300]}")
                        except Exception as e:
                            logger.warning(f"Subprocess error for page {page_num}: {e}")

                    if not converted and os.path.exists(local_pdf):
                        try:
                            logger.info(f"Trying pdf2docx fallback for page {page_num}...")
                            from pdf2docx import Converter
                            cv = Converter(local_pdf)
                            cv.convert(chunk_docx)
                            cv.close()
                            if os.path.exists(chunk_docx) and os.path.getsize(chunk_docx) > 0:
                                converted = True
                        except Exception as fe:
                            logger.warning(f"pdf2docx fallback failed for page {page_num}: {fe}")

                    try:
                        if os.path.exists(temp_chunk_ocr): os.remove(temp_chunk_ocr)
                    except Exception: pass

                # Clean watermark and collect path
                if converted and os.path.exists(chunk_docx):
                    try:
                        self._clean_docx(chunk_docx, chunk_docx, is_arabic)
                    except Exception as ce:
                        logger.warning(f"Failed to clean watermark for page {page_num}: {ce}")
                    
                    with completed_lock:
                        converted_results[idx] = chunk_docx
                        if is_gemini_success:
                            gemini_success_count += 1

                # Cleanup single page PDF immediately
                try:
                    if os.path.exists(chunk_pdf): os.remove(chunk_pdf)
                except Exception: pass

                # Update progress
                with completed_lock:
                    completed_count += 1
                    if progress_callback:
                        try: progress_callback(completed_count, total_pages)
                        except Exception: pass

            # Run parallel execution using ThreadPoolExecutor
            num_workers = min(total_pages, len(api_key_groups) if api_key_groups else 4)
            num_workers = min(num_workers, 8) # Limit workers to 8 to avoid resource starvation
            logger.info(f"Starting parallel page conversion with {num_workers} threads...")
            
            with ThreadPoolExecutor(max_workers=num_workers) as executor:
                executor.map(convert_page_worker, range(total_pages))

            doc_src.close()

            # Gather converted files in sorted order
            converted_docx_files = [converted_results[i] for i in sorted(converted_results.keys())]

            # Merge converted page files into the single final output_path
            if converted_docx_files:
                logger.info(f"Merging {len(converted_docx_files)} page files into a single DOCX...")
                try:
                    from docxcompose.composer import Composer
                    from docx import Document
                    
                    master_doc = Document(converted_docx_files[0])
                    composer = Composer(master_doc)
                    
                    for chunk_file in converted_docx_files[1:]:
                        if os.path.exists(chunk_file) and os.path.getsize(chunk_file) > 0:
                            master_doc.add_page_break()
                            composer.append(Document(chunk_file))
                            
                    composer.save(output_path)
                    logger.info("Successfully merged all pages into a single DOCX file.")
                except Exception as merge_err:
                    logger.error(f"Failed to merge DOCX files: {merge_err}", exc_info=True)
                    # If merge fails, copy the first successfully converted page as output fallback
                    import shutil
                    shutil.copy(converted_docx_files[0], output_path)

                # Cleanup intermediate docx files
                for chunk_file in converted_docx_files:
                    try:
                        if os.path.exists(chunk_file):
                            os.remove(chunk_file)
                    except Exception:
                        pass
            else:
                raise RuntimeError("Failed to convert any of the pages.")

            return {
                'pages': total_pages,
                'is_scanned': is_scanned,
                'time': round(time.time() - start_time, 2),
                'engine': 'gemini_vision' if gemini_success_count > 0 else 'local',
                'language': 'ar' if is_arabic else 'en',
                'is_rtl': is_arabic
            }
        # --------------------------------------

        
        gemini_error = None
        
        # Already determined use_ai at start of method
        pass
                
        if use_ai:
            logger.info(f"Using Gemini Vision Layout Reconstruction API (mode={mode})...")
            try:
                self._convert_via_gemini_vision(input_path, output_path, language, gemini_api_key, progress_callback=progress_callback)
                return {
                    'pages': total_pages,
                    'is_scanned': True,
                    'time': round(time.time() - start_time, 2),
                    'engine': 'gemini_vision',
                    'language': 'ar' if is_arabic else 'en',
                    'is_rtl': is_arabic
                }
            except Exception as e:
                import traceback
                traceback.print_exc()
                gemini_error = str(e)
                logger.error(f"Gemini Vision pipeline failed: {gemini_error}. Falling back to local engine.")
        
        # Step 1: Pre-process PDF if scanned
        working_pdf = input_path
        temp_pdf = input_path.replace('.pdf', '_temp_searchable.pdf')
        temp_flattened_pdf = input_path.replace('.pdf', '_flattened.pdf')
        temp_docx = output_path.replace('.docx', '_raw.docx')
        
        if is_scanned or mode == 'ocr':
            # For 1-page scanned PDFs: use Tesseract + python-docx directly
            logger.info("Scanned document detected (single page). Running OCR to extract text directly to DOCX...")
            try:
                import pytesseract
                from docx import Document as DocxDocument
                from docx.shared import Pt
                from docx.enum.text import WD_ALIGN_PARAGRAPH
                from PIL import Image, ImageEnhance, ImageFilter
                import io

                page_doc_ocr = fitz.open(input_path)
                pix = page_doc_ocr[0].get_pixmap(matrix=fitz.Matrix(300/72, 300/72))
                img_ocr = Image.open(io.BytesIO(pix.tobytes('png')))
                page_doc_ocr.close()

                img_gray = img_ocr.convert('L')
                img_gray = ImageEnhance.Contrast(img_gray).enhance(1.8)
                img_gray = img_gray.filter(ImageFilter.SHARPEN)

                tess_lang = 'ara' if language in ('auto', 'ar', '') else 'eng'
                ocr_text = pytesseract.image_to_string(img_gray, lang=tess_lang, config='--oem 3 --psm 4').strip()

                word_doc = DocxDocument()
                is_rtl_page = language in ('auto', 'ar', '')
                if ocr_text:
                    for line in ocr_text.split('\n'):
                        para = word_doc.add_paragraph(line)
                        para.alignment = WD_ALIGN_PARAGRAPH.RIGHT if is_rtl_page else WD_ALIGN_PARAGRAPH.LEFT
                        for run in para.runs:
                            run.font.size = Pt(12)
                else:
                    word_doc.add_paragraph('')

                word_doc.save(temp_docx)
                logger.info(f"Single-page OCR DOCX created successfully ({len(ocr_text)} chars)")
            except Exception as oe:
                logger.error(f"Single page OCR failed: {oe}")
                raise RuntimeError(f"OCR failed for single page: {oe}")
        else:
            # Step 2: Convert to DOCX using Spire (Commercial Layout Engine) or pdf2docx
            logger.info(f"Converting {total_pages} pages to DOCX (is_arabic={is_arabic})...")
            
            def try_spire():
                spire_doc = PdfDocument()
                spire_doc.LoadFromFile(working_pdf)
                spire_doc.SaveToFile(temp_docx, FileFormat.DOCX)
                spire_doc.Close()
                
            def try_pdf2docx():
                from pdf2docx import Converter
                cv = Converter(working_pdf)
                cv.convert(temp_docx)
                cv.close()

            try:
                if is_arabic:
                    try:
                        try_pdf2docx()
                    except Exception as e:
                        logger.warning(f"pdf2docx failed: {e}. Trying Spire fallback...")
                        try_spire()
                else:
                    try:
                        try_spire()
                    except Exception as e:
                        logger.warning(f"Local Spire engine failed: {e}. Trying pdf2docx fallback...")
                        try_pdf2docx()
            except Exception as final_err:
                logger.error(f"Both local engines failed: {final_err}")
                raise RuntimeError(f"Engine failed to parse layout. Detail: {final_err}")

        # Step 3: Post-process (Remove Watermarks)
        logger.info("Cleaning up document layout and watermarks...")
        self._clean_docx(temp_docx, output_path, is_arabic)

        # Cleanup temp files
        if os.path.exists(temp_pdf): os.remove(temp_pdf)
        if os.path.exists(temp_docx): os.remove(temp_docx)
        if os.path.exists(temp_flattened_pdf):
            try: os.remove(temp_flattened_pdf)
            except Exception: pass

        return {
            'pages': total_pages,
            'is_scanned': is_scanned,
            'time': round(time.time() - start_time, 2),
            'engine': f"local: {gemini_error}" if gemini_error else "local",
            'language': 'ar' if is_arabic else 'en',
            'is_rtl': is_arabic
        }

    # ─────────────────────────────────────────────
    # INTERNAL METHODS
    # ─────────────────────────────────────────────

    def _is_scanned(self, doc: fitz.Document) -> bool:
        """Heuristics to determine if PDF is a scanned image.
        Checks ALL pages - if ANY page has real text it's a digital PDF."""
        if len(doc) == 0: return False
        
        # Check every page. If ANY page has meaningful text → digital PDF
        for i in range(len(doc)):
            text = doc[i].get_text('text').strip()
            if len(text) > 50:
                return False  # Found real text anywhere → definitely digital
                
        return True  # No page had any text → scanned

    def _make_searchable_pdf(self, input_path: str, output_path: str, lang: str):
        """Uses Tesseract to create a searchable PDF with exact layout retention."""
        doc = fitz.open(input_path)
        pdf_bytes_list = []
        
        # Decide tesseract language
        tess_lang = 'ara' if lang in ('auto', 'ar', '') else 'eng'
        
        for i in range(len(doc)):
            page = doc[i]
            # Render at 300 DPI for high quality OCR
            pix = page.get_pixmap(matrix=fitz.Matrix(300/72, 300/72))
            img = Image.open(io.BytesIO(pix.tobytes('png')))
            
            # Enhancements for better OCR accuracy
            img_gray = img.convert('L')
            img_gray = ImageEnhance.Contrast(img_gray).enhance(1.8)
            img_gray = img_gray.filter(ImageFilter.SHARPEN)
            
            # Tesseract generates a perfectly aligned PDF page
            try:
                page_pdf = pytesseract.image_to_pdf_or_hocr(
                    img_gray, lang=tess_lang, extension='pdf', config='--oem 3 --psm 4'
                )
                pdf_bytes_list.append(page_pdf)
            except Exception as e:
                logger.warning(f"OCR failed on page {i}: {e}")
        
        doc.close()
        
        # Merge all generated PDF pages into one
        merged_doc = fitz.open()
        for pdf_b in pdf_bytes_list:
            page_doc = fitz.open("pdf", pdf_b)
            merged_doc.insert_pdf(page_doc)
            page_doc.close()
            
        merged_doc.save(output_path)
        merged_doc.close()

    def _fix_spire_docx_styles(self, docx_path: str):
        """Fixes 'WD_STYLE_TYPE has no XML mapping for list' error caused by Spire."""
        import zipfile
        import os
        import tempfile
        import shutil
        
        needs_fix = False
        try:
            with zipfile.ZipFile(docx_path, 'r') as zin:
                if 'word/styles.xml' in zin.namelist():
                    styles_content = zin.read('word/styles.xml')
                    if b'w:type="list"' in styles_content:
                        needs_fix = True
        except Exception:
            return

        if not needs_fix:
            return

        temp_fd, temp_path = tempfile.mkstemp(suffix='.docx')
        os.close(temp_fd)

        with zipfile.ZipFile(docx_path, 'r') as zin:
            with zipfile.ZipFile(temp_path, 'w') as zout:
                for item in zin.infolist():
                    buffer = zin.read(item.filename)
                    if item.filename == 'word/styles.xml':
                        buffer = buffer.replace(b'w:type="list"', b'w:type="paragraph"')
                    zout.writestr(item, buffer)
                    
        shutil.move(temp_path, docx_path)

    def _clean_docx(self, input_docx: str, output_docx: str, is_arabic: bool = False):
        """Removes Spire.Pdf 'Evaluation Warning' watermarks seamlessly and forces RTL alignment for Arabic."""
        self._fix_spire_docx_styles(input_docx)
        
        # 1. Deep clean using ZIP and lxml with heuristic fragment detection
        import zipfile
        import io
        from lxml import etree
        import re
        
        temp_buffer = io.BytesIO()
        with zipfile.ZipFile(input_docx, 'r') as zin:
            with zipfile.ZipFile(temp_buffer, 'w') as zout:
                for item in zin.infolist():
                    data = zin.read(item.filename)
                    if item.filename.startswith('word/') and item.filename.endswith('.xml'):
                        try:
                            tree = etree.fromstring(data)
                            namespaces = {'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'}
                            modified = False
                            
                            # Pass 1: Standard detection (single paragraph contains watermark)
                            for p in tree.xpath('.//w:p', namespaces=namespaces):
                                text = ''.join(node.text for node in p.xpath('.//w:t', namespaces=namespaces) if node.text)
                                normalized = re.sub(r'[^a-zA-Z]', '', text).lower()
                                if 'evaluationwarning' in normalized or 'spirepdf' in normalized:
                                    p.getparent().remove(p)
                                    modified = True
                            
                            # Pass 2: Fragmented detection (Arabic mode obfuscation)
                            # Spire splits the text into many red floating paragraphs
                            red_floating_ps = []
                            for p in tree.xpath('.//w:p', namespaces=namespaces):
                                has_red = False
                                for color in p.xpath('.//w:color', namespaces=namespaces):
                                    val = color.get('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}val')
                                    if val and val.lower() == 'ff0000':
                                        has_red = True
                                        break
                                if has_red and p.xpath('.//w:framePr', namespaces=namespaces):
                                    red_floating_ps.append(p)
                                    
                            if red_floating_ps:
                                combined_text = ""
                                for p in red_floating_ps:
                                    text = ''.join(node.text for node in p.xpath('.//w:t', namespaces=namespaces) if node.text)
                                    combined_text += text
                                normalized = re.sub(r'[^a-zA-Z]', '', combined_text).lower()
                                if 'evaluationwarning' in normalized or 'spirepdf' in normalized:
                                    for p in red_floating_ps:
                                        p.getparent().remove(p)
                                    modified = True
                                    
                            if modified:
                                data = etree.tostring(tree, encoding='utf-8', xml_declaration=True)
                        except Exception as e:
                            pass
                    zout.writestr(item, data)
                    
        # Write back to disk
        with open(input_docx, 'wb') as f:
            f.write(temp_buffer.getvalue())

        # 2. Use python-docx for RTL formatting
        from docx import Document
        doc = Document(input_docx)
        from docx.enum.text import WD_ALIGN_PARAGRAPH
        from docx.oxml import OxmlElement
        from docx.oxml.ns import qn
        
        def _apply_rtl_paragraph(p):
            """Apply RTL formatting correctly: w:bidi MUST come before w:jc in pPr per OOXML spec."""
            # Set alignment via python-docx (sets w:jc)
            import re as _re
            txt = p.text.strip()
            if p.alignment == WD_ALIGN_PARAGRAPH.CENTER:
                is_list_or_body = (len(txt) > 30 or bool(_re.search(r'^[\d\u0660-\u0669][\s\-\.\)]+', txt)))
                if not is_list_or_body:
                    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                else:
                    p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
            else:
                p.alignment = WD_ALIGN_PARAGRAPH.RIGHT

            pPr = p._p.get_or_add_pPr()
            # Remove any existing w:bidi to avoid duplicates
            for old in pPr.findall(qn('w:bidi')):
                pPr.remove(old)
            # Create w:bidi element with explicit val="1"
            bidi_el = OxmlElement('w:bidi')
            bidi_el.set(qn('w:val'), '1')
            # CRITICAL: insert w:bidi BEFORE w:jc (and other tags that come after it in the schema)
            # OOXML schema order: pStyle > numPr > bidi > jc > ...
            after_bidi_tags = [
                'w:jc', 'w:textDirection', 'w:textAlignment', 'w:textboxTightWrap',
                'w:outlineLvl', 'w:divId', 'w:cnfStyle', 'w:rPr', 'w:sectPr', 'w:pPrChange'
            ]
            inserted = False
            for i, child in enumerate(pPr):
                if child.tag in [qn(tag) for tag in after_bidi_tags]:
                    pPr.insert(i, bidi_el)
                    inserted = True
                    break
            if not inserted:
                pPr.append(bidi_el)

            for run in p.runs:
                rPr = run._r.get_or_add_rPr()
                if not rPr.findall(qn('w:rtl')):
                    rPr.append(OxmlElement('w:rtl'))
                if not rPr.findall(qn('w:lang')):
                    lang_el = OxmlElement('w:lang')
                    lang_el.set(qn('w:val'), 'ar-SA')
                    lang_el.set(qn('w:bidi'), 'ar-SA')
                    rPr.append(lang_el)
                if run.font.size and not rPr.findall(qn('w:szCs')):
                    szCs = OxmlElement('w:szCs')
                    szCs.set(qn('w:val'), str(int(run.font.size.pt * 2)))
                    rPr.append(szCs)

        for p in doc.paragraphs:
            if is_arabic:
                _apply_rtl_paragraph(p)

        if is_arabic:
            for table in doc.tables:
                for row in table.rows:
                    for cell in row.cells:
                        for p in cell.paragraphs:
                            _apply_rtl_paragraph(p)
        doc.save(output_docx)

        # Apply ZIP patch to enforce custom style, bidi, and jc ordering for all Arabic documents
        if is_arabic:
            try:
                import io as _io
                import zipfile as _zf
                from lxml import etree as _ET
                
                _W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
                def _qw(tag):
                    return f"{{{_W}}}{tag}"
                
                _AFTER_BIDI = {
                    _qw('adjustRightInd'), _qw('snapToGrid'), _qw('spacing'), _qw('ind'),
                    _qw('contextualSpacing'), _qw('jc'), _qw('textDirection'), _qw('textAlignment'),
                    _qw('textboxTightWrap'), _qw('outlineLvl'), _qw('divId'), _qw('cnfStyle'),
                    _qw('rPr'), _qw('sectPr'), _qw('pPrChange')
                }
                
                _AFTER_JC = {
                    _qw('textDirection'), _qw('textAlignment'), _qw('textboxTightWrap'),
                    _qw('outlineLvl'), _qw('divId'), _qw('cnfStyle'), _qw('rPr'), _qw('sectPr'), _qw('pPrChange')
                }

                def _patch_pPr_base(pPr):
                    for old in pPr.findall(_qw('bidi')):
                        pPr.remove(old)
                    for old in pPr.findall(_qw('jc')):
                        pPr.remove(old)
                    bidi_el = _ET.Element(_qw('bidi'))
                    bidi_el.set(_qw('val'), '1')
                    pos_bidi = next((i for i, c in enumerate(pPr) if c.tag in _AFTER_BIDI), len(pPr))
                    pPr.insert(pos_bidi, bidi_el)
                    jc_el = _ET.Element(_qw('jc'))
                    jc_el.set(_qw('val'), 'start')
                    pos_jc = next((i for i, c in enumerate(pPr) if c.tag in _AFTER_JC), len(pPr))
                    pPr.insert(pos_jc, jc_el)

                def _patch_paragraph_pPr(pPr):
                    for old in pPr.findall(_qw('pStyle')):
                        pPr.remove(old)
                    pStyle = _ET.Element(_qw('pStyle'))
                    pStyle.set(_qw('val'), 'ArabicBody')
                    pPr.insert(0, pStyle)
                    _patch_pPr_base(pPr)

                def _patch_rPr(rPr):
                    if rPr.find(_qw('rtl')) is None:
                        _ET.SubElement(rPr, _qw('rtl'))
                    for old in rPr.findall(_qw('lang')):
                        rPr.remove(old)
                    lang_el = _ET.SubElement(rPr, _qw('lang'))
                    lang_el.set(_qw('val'), 'ar-SA')
                    lang_el.set(_qw('bidi'), 'ar-SA')

                buf = _io.BytesIO()
                with _zf.ZipFile(output_docx, 'r') as zin, _zf.ZipFile(buf, 'w', _zf.ZIP_DEFLATED) as zout:
                    for item in zin.infolist():
                        data = zin.read(item.filename)

                        if item.filename == 'word/document.xml':
                            try:
                                tree = _ET.fromstring(data)
                                body = tree.find(f'{{{_W}}}body')
                                for p_el in body.findall(f'.//{_qw("p")}'):
                                    pPr = p_el.find(_qw('pPr'))
                                    if pPr is None:
                                        pPr = _ET.Element(_qw('pPr'))
                                        p_el.insert(0, pPr)
                                    _patch_paragraph_pPr(pPr)
                                    for r_el in p_el.findall(_qw('r')):
                                        rPr = r_el.find(_qw('rPr'))
                                        if rPr is None:
                                            rPr = _ET.Element(_qw('rPr'))
                                            r_el.insert(0, rPr)
                                        _patch_rPr(rPr)
                                sectPr = body.find(_qw('sectPr'))
                                if sectPr is not None:
                                    for old in sectPr.findall(_qw('bidi')):
                                        sectPr.remove(old)
                                    bidi_sect = _ET.SubElement(sectPr, _qw('bidi'))
                                    bidi_sect.set(_qw('val'), '1')
                                data = _ET.tostring(tree, xml_declaration=True, encoding='UTF-8', standalone=True)
                            except Exception as _e:
                                logger.warning(f'RTL patch document.xml in clean_docx: {_e}')

                        elif item.filename == 'word/styles.xml':
                            try:
                                tree = _ET.fromstring(data)
                                
                                # Check if ArabicBody style already exists in styles.xml
                                style_exists = False
                                for st in tree.findall(_qw('style')):
                                    if st.get(_qw('styleId')) == 'ArabicBody':
                                        style_exists = True
                                        break
                                
                                if not style_exists:
                                    # Create ArabicBody style element dynamically
                                    arabic_body = _ET.Element(_qw('style'))
                                    arabic_body.set(_qw('type'), 'paragraph')
                                    arabic_body.set(_qw('styleId'), 'ArabicBody')
                                    
                                    _ET.SubElement(arabic_body, _qw('name')).set(_qw('val'), 'ArabicBody')
                                    _ET.SubElement(arabic_body, _qw('basedOn')).set(_qw('val'), 'Normal')
                                    _ET.SubElement(arabic_body, _qw('qFormat'))
                                    
                                    pPr_ab = _ET.SubElement(arabic_body, _qw('pPr'))
                                    bidi_ab = _ET.SubElement(pPr_ab, _qw('bidi'))
                                    bidi_ab.set(_qw('val'), '1')
                                    
                                    sp_ab = _ET.SubElement(pPr_ab, _qw('spacing'))
                                    sp_ab.set(_qw('after'), '40')
                                    sp_ab.set(_qw('before'), '0')
                                    sp_ab.set(_qw('line'), '276')
                                    sp_ab.set(_qw('lineRule'), 'auto')
                                    
                                    jc_ab = _ET.SubElement(pPr_ab, _qw('jc'))
                                    jc_ab.set(_qw('val'), 'start')
                                    
                                    rPr_ab = _ET.SubElement(arabic_body, _qw('rPr'))
                                    fonts_ab = _ET.SubElement(rPr_ab, _qw('rFonts'))
                                    fonts_ab.set(_qw('ascii'), 'Traditional Arabic')
                                    fonts_ab.set(_qw('hAnsi'), 'Traditional Arabic')
                                    fonts_ab.set(_qw('cs'), 'Traditional Arabic')
                                    
                                    sz_ab = _ET.SubElement(rPr_ab, _qw('sz'))
                                    sz_ab.set(_qw('val'), '23')
                                    szCs_ab = _ET.SubElement(rPr_ab, _qw('szCs'))
                                    szCs_ab.set(_qw('val'), '23')
                                    
                                    _ET.SubElement(rPr_ab, _qw('rtl'))
                                    
                                    lang_ab = _ET.SubElement(rPr_ab, _qw('lang'))
                                    lang_ab.set(_qw('val'), 'ar-SA')
                                    lang_ab.set(_qw('bidi'), 'ar-SA')
                                    
                                    tree.append(arabic_body)
                                    logger.info('Dynamically injected ArabicBody style into styles.xml')

                                # Fix other style-level pPr definitions
                                for pPr in tree.findall(f'.//{_qw("pPr")}'):
                                    _patch_pPr_base(pPr)
                                
                                data = _ET.tostring(tree, xml_declaration=True, encoding='UTF-8', standalone=True)
                            except Exception as _e:
                                logger.warning(f'RTL patch styles.xml in clean_docx: {_e}')

                        zout.writestr(item, data)

                with open(output_docx, 'wb') as _f:
                    _f.write(buf.getvalue())
                logger.info('_clean_docx: Definitive Arabic RTL ZIP patch applied successfully.')
            except Exception as _rtl_err:
                logger.warning(f'_clean_docx: Definitive Arabic RTL patch failed: {_rtl_err}')

    def _convert_via_gemini_vision(self, input_path: str, output_path: str, lang: str, api_key: str, progress_callback=None):
        import fitz
        import base64
        import requests
        import os
        import cv2
        import numpy as np
        from docx import Document
        from docx.shared import Pt, RGBColor
        from docx.enum.text import WD_ALIGN_PARAGRAPH
        from docx.oxml import parse_xml, OxmlElement
        from docx.oxml.ns import nsdecls, qn
        from bs4 import BeautifulSoup
        import re
        
        def extract_page_visual_elements(page, unique_prefix, page_idx, output_path, is_book_or_doc):
            visual_elements = []
            image_list = page.get_images(full=True)
            
            is_scanned = False
            # 1. Page Count/Document Type Safety Filter:
            # For books/documents, we completely disable the scanned visual element crop engine to protect beautiful decorative borders, frames, and margins.
            if not is_book_or_doc:
                if len(image_list) == 1:
                    xref = image_list[0][0]
                    rects = page.get_image_rects(xref)
                    if rects:
                        r = rects[0]
                        page_area = page.rect.width * page.rect.height
                        img_area = (r.x1 - r.x0) * (r.y1 - r.y0)
                        if img_area / page_area > 0.8:
                            is_scanned = True
                        
            if is_scanned:
                try:
                    pix = page.get_pixmap(matrix=fitz.Matrix(2, 2))
                    img_data = np.frombuffer(pix.samples, dtype=np.uint8).reshape((pix.h, pix.w, 3))
                    img = cv2.cvtColor(img_data, cv2.COLOR_RGB2BGR)
                    
                    h_img, w_img, _ = img.shape
                    gray = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)
                    
                    # Globally robust adaptive thresholding to handle any contrast or background lighting
                    thresh = cv2.adaptiveThreshold(gray, 255, cv2.ADAPTIVE_THRESH_GAUSSIAN_C, cv2.THRESH_BINARY_INV, 11, 2)
                    contours, hierarchy = cv2.findContours(thresh, cv2.RETR_TREE, cv2.CHAIN_APPROX_SIMPLE)
                    
                    qr_idx = 0
                    bc_idx = 0
                    logo_idx = 0
                    
                    # 1. ROBUST QR CODE DETECTION USING FINDER-PATTERN GROUPING
                    # Group small square contours (nested finding patterns) within 250px center-to-center distance
                    # to reliably extract the QR code even when watermarks or signatures touch/corrupt its boundaries.
                    finder_patterns = []
                    for idx, c in enumerate(contours):
                        x, y, w, h = cv2.boundingRect(c)
                        aspect_ratio = w / float(h) if h > 0 else 0
                        if (12 < w < 85 and 12 < h < 85 and 0.75 < aspect_ratio < 1.3):
                            nested_count = 0
                            if hierarchy is not None and idx < len(hierarchy[0]):
                                child_idx = hierarchy[0][idx][2]
                                while child_idx != -1:
                                    nested_count += 1
                                    if child_idx < len(hierarchy[0]):
                                        child_idx = hierarchy[0][child_idx][2]
                                    else:
                                        break
                            if nested_count >= 2:  # Requires 2 levels of nesting to filter out plain text/noise!
                                finder_patterns.append((x, y, w, h, x + w/2.0, y + h/2.0))
                                
                    used_finders = set()
                    qr_boxes = []
                    for i, f1 in enumerate(finder_patterns):
                        if i in used_finders:
                            continue
                        group = [f1]
                        for j, f2 in enumerate(finder_patterns):
                            if i != j and j not in used_finders:
                                dist = np.sqrt((f1[4] - f2[4])**2 + (f1[5] - f2[5])**2)
                                if dist < 250:
                                    group.append(f2)
                        if len(group) >= 3:  # Requires at least 3 finder patterns to form a true QR code!
                            for f in group:
                                try:
                                    used_finders.add(finder_patterns.index(f))
                                except:
                                    pass
                            min_x = min(f[0] for f in group)
                            min_y = min(f[1] for f in group)
                            max_x = max(f[0] + f[2] for f in group)
                            max_y = max(f[1] + f[3] for f in group)
                            
                            qw = max_x - min_x
                            qh = max_y - min_y
                            pad_x = int(qw * 0.35)
                            pad_y = int(qh * 0.35)
                            
                            x_qr = max(0, min_x - pad_x)
                            y_qr = max(0, min_y - pad_y)
                            w_qr = min(w_img - x_qr, qw + 2 * pad_x)
                            h_qr = min(h_img - y_qr, qh + 2 * pad_y)
                            
                            crop_qr = img[y_qr:y_qr+h_qr, x_qr:x_qr+w_qr]
                            temp_img_name = f"temp_crop_qr_{unique_prefix}_{page_idx}_{qr_idx}.png"
                            temp_img_path = os.path.join(os.path.dirname(output_path), temp_img_name)
                            cv2.imwrite(temp_img_path, crop_qr)
                            
                            visual_elements.append({
                                'path': temp_img_path,
                                'y_rel': (y_qr + h_qr/2.0) / h_img,
                                'x_rel': (x_qr + w_qr/2.0) / w_img,
                                'type': 'qr',
                                'x': x_qr, 'y': y_qr, 'w': w_qr, 'h': h_qr
                            })
                            qr_idx += 1
                            qr_boxes.append((x_qr, y_qr, w_qr, h_qr))
                            
                    # 2. MAIN CONTOUR SCAN (BARCODES, LOGOS, FALLBACK QRS)
                    for idx, c in enumerate(contours):
                        x, y, w, h = cv2.boundingRect(c)
                        
                        # Skip if center lies inside any already detected QR box to prevent duplicates
                        is_inside_qr = False
                        cx_c, cy_c = x + w/2.0, y + h/2.0
                        for (qx, qy, qw, qh) in qr_boxes:
                            if qx <= cx_c <= qx + qw and qy <= cy_c <= qy + qh:
                                is_inside_qr = True
                                break
                        if is_inside_qr:
                            continue
                            
                        # Skip noise, extremely small icons, or page-covering background shapes
                        if w < 40 or h < 15 or w > w_img * 0.85 or h > h_img * 0.85:
                            continue
                            
                        aspect_ratio = w / float(h) if h > 0 else 0
                        y_rel = (y + h/2.0) / h_img
                        x_rel = (x + w/2.0) / w_img
                        
                        crop = img[y:y+h, x:x+w]
                        crop_gray = gray[y:y+h, x:x+w]
                        
                        edges = cv2.Canny(crop_gray, 50, 150)
                        edge_density = np.sum(edges > 0) / float(edges.size) if edges.size > 0 else 0
                        
                        grad_x = cv2.Sobel(crop_gray, cv2.CV_32F, 1, 0, ksize=3)
                        grad_y = cv2.Sobel(crop_gray, cv2.CV_32F, 0, 1, ksize=3)
                        mean_grad_x = np.mean(np.abs(grad_x)) if grad_x.size > 0 else 0
                        mean_grad_y = np.mean(np.abs(grad_y)) if grad_y.size > 0 else 0
                        gradient_ratio = mean_grad_x / (mean_grad_y + 1e-5)
                        
                        # A. Fallback QR detection (single closed contour QR codes) - skipped in header (y_rel < 0.28)
                        if (y_rel >= 0.28 and
                            0.8 < aspect_ratio < 1.25 and 
                            50 < w < 450 and 50 < h < 450 and 
                            edge_density > 0.08 and 
                            0.5 < gradient_ratio < 2.0):
                            
                            temp_img_name = f"temp_crop_qr_{unique_prefix}_{page_idx}_{qr_idx}.png"
                            temp_img_path = os.path.join(os.path.dirname(output_path), temp_img_name)
                            cv2.imwrite(temp_img_path, crop)
                            
                            visual_elements.append({
                                'path': temp_img_path,
                                'y_rel': y_rel,
                                'x_rel': x_rel,
                                'type': 'qr',
                                'area_ratio': (w * h) / (w_img * h_img),
                                'x': x, 'y': y, 'w': w, 'h': h
                            })
                            qr_idx += 1
                            
                        # B. Global Barcode detection
                        elif (y_rel < 0.28 and
                              2.0 < aspect_ratio < 9.0 and 
                              50 < w < 500 and 12 < h < 120 and 
                              edge_density > 0.05 and 
                              gradient_ratio > 1.2):
                              
                            temp_img_name = f"temp_crop_bc_{unique_prefix}_{page_idx}_{bc_idx}.png"
                            temp_img_path = os.path.join(os.path.dirname(output_path), temp_img_name)
                            cv2.imwrite(temp_img_path, crop)
                            
                            visual_elements.append({
                                'path': temp_img_path,
                                'y_rel': y_rel,
                                'x_rel': x_rel,
                                'type': 'barcode',
                                'area_ratio': (w * h) / (w_img * h_img),
                                'x': x, 'y': y, 'w': w, 'h': h
                            })
                            bc_idx += 1
                            
                        # C. Global Logo/Emblem detection in the header (like Coat of Arms or corporate logos)
                        elif (y_rel < 0.28 and
                              0.65 < aspect_ratio < 1.6 and 
                              45 < w < 280 and 45 < h < 280 and 
                              edge_density > 0.05):
                              
                            temp_img_name = f"temp_crop_logo_{unique_prefix}_{page_idx}_{logo_idx}.png"
                            temp_img_path = os.path.join(os.path.dirname(output_path), temp_img_name)
                            cv2.imwrite(temp_img_path, crop)
                            
                            visual_elements.append({
                                'path': temp_img_path,
                                'y_rel': y_rel,
                                'x_rel': x_rel,
                                'type': 'logo',
                                'area_ratio': (w * h) / (w_img * h_img),
                                'x': x, 'y': y, 'w': w, 'h': h
                            })
                            logo_idx += 1
                            
                    # Bounding Box Deduplication (NMS)
                    deduped_elements = []
                    visual_elements.sort(key=lambda elem: elem.get('w', 0) * elem.get('h', 0), reverse=True)
                    for elem in visual_elements:
                        x1, y1, w1, h1 = elem.get('x', 0), elem.get('y', 0), elem.get('w', 0), elem.get('h', 0)
                        area1 = w1 * h1
                        overlap = False
                        for accepted in deduped_elements:
                            x2, y2, w2, h2 = accepted.get('x', 0), accepted.get('y', 0), accepted.get('w', 0), accepted.get('h', 0)
                            # Intersection
                            ix = max(x1, x2)
                            iy = max(y1, y2)
                            iw = min(x1 + w1, x2 + w2) - ix
                            ih = min(y1 + h1, y2 + h2) - iy
                            if iw > 0 and ih > 0:
                                intersection_area = iw * ih
                                smaller_area = min(area1, w2 * h2)
                                if intersection_area / smaller_area > 0.4:
                                    overlap = True
                                    break
                        if not overlap:
                            deduped_elements.append(elem)
                        else:
                            # Delete overlapping cropped image to clean disk
                            try:
                                if os.path.exists(elem['path']):
                                    os.remove(elem['path'])
                            except:
                                pass
                    visual_elements = deduped_elements
                            
                    # 3. Candidate Count Safety Filter (Relaxed to > 3 to allow barcode, QR, and logo together)
                    if len(visual_elements) > 3:
                        logger.info(f"Page {page_idx+1} has {len(visual_elements)} visual elements. This is likely a decorative layout or border, not a receipt. Skipping all crops.")
                        for elem in visual_elements:
                            try:
                                os.remove(elem['path'])
                            except:
                                pass
                        visual_elements = []
                except Exception as se:
                    logger.error(f"Failed OpenCV scan on page {page_idx+1}: {se}")
                    
            if not visual_elements and image_list:
                img_idx = 0
                if not is_arabic:
                    # English/LTR files: merge overlapping/adjacent image rects and render page clips
                    raw_rects = []
                    for img_info in image_list:
                        xref = img_info[0]
                        rects = page.get_image_rects(xref)
                        for r in rects:
                            page_area = page.rect.width * page.rect.height
                            img_area = (r.x1 - r.x0) * (r.y1 - r.y0)
                            if img_area / page_area > 0.8:
                                continue
                            w = r.x1 - r.x0
                            h = r.y1 - r.y0
                            # Enforce a size threshold of 10pt to completely skip tiny pattern tiles and page border shapes
                            if w < 10 or h < 10:
                                continue
                            raw_rects.append(r)
                            
                    # Merge overlapping/adjacent rectangles
                    merged_rects = []
                    for r in raw_rects:
                        merged_any = False
                        for idx, m in enumerate(merged_rects):
                            # Inflate m by a threshold of 12pt to check if r is close/overlapping
                            inflated_m = fitz.Rect(m.x0 - 12.0, m.y0 - 12.0, m.x1 + 12.0, m.y1 + 12.0)
                            intersect = inflated_m & r
                            if not intersect.is_empty:
                                merged_rects[idx] = m | r
                                merged_any = True
                                break
                        if not merged_any:
                            merged_rects.append(r)
                            
                    # Run a cascading merge pass to combine newly overlapping groups
                    changed = True
                    while changed:
                        changed = False
                        new_merged = []
                        for r in merged_rects:
                            merged_any = False
                            for idx, m in enumerate(new_merged):
                                inflated_m = fitz.Rect(m.x0 - 12.0, m.y0 - 12.0, m.x1 + 12.0, m.y1 + 12.0)
                                intersect = inflated_m & r
                                if not intersect.is_empty:
                                    new_merged[idx] = m | r
                                    merged_any = True
                                    changed = True
                                    break
                            if not merged_any:
                                new_merged.append(r)
                        merged_rects = new_merged
                        
                    for r in merged_rects:
                        try:
                            # Safely clip within page boundary
                            r_clip = r.intersect(page.rect)
                            if r_clip.is_empty or (r_clip.x1 - r_clip.x0) < 10 or (r_clip.y1 - r_clip.y0) < 10:
                                continue
                            
                            pix = page.get_pixmap(clip=r_clip, matrix=fitz.Matrix(220/72, 220/72))
                            image_bytes = pix.tobytes("png")
                            
                            temp_img_name = f"temp_ext_{unique_prefix}_{page_idx}_{img_idx}.png"
                            temp_img_path = os.path.join(os.path.dirname(output_path), temp_img_name)
                            with open(temp_img_path, "wb") as f:
                                f.write(image_bytes)
                                
                            visual_elements.append({
                                'path': temp_img_path,
                                'y_rel': ((r_clip.y0 + r_clip.y1) / 2.0) / page.rect.height,
                                'x_rel': ((r_clip.x0 + r_clip.x1) / 2.0) / page.rect.width,
                                'type': 'extracted',
                                'area_ratio': ((r_clip.x1 - r_clip.x0) * (r_clip.y1 - r_clip.y0)) / (page.rect.width * page.rect.height)
                            })
                            img_idx += 1
                        except Exception as ee:
                            logger.error(f"Failed standard clip extract on page {page_idx+1}: {ee}")
                else:
                    # Arabic files: keep raw image bytes extraction for invoice safety
                    # For BOOKS/DOCUMENTS, use a STRICTER threshold to avoid extracting
                    # decorative frames/borders that cover large page areas.
                    # Receipts (non-book): skip only full-page images (> 80%)
                    # Books/Docs:          skip anything covering > 40% (frames, borders)
                    ar_img_skip_thresh = 0.40 if is_book_or_doc else 0.80
                    
                    for img_info in image_list:
                        try:
                            xref = img_info[0]
                            rects = page.get_image_rects(xref)
                            if rects:
                                r = rects[0]
                                page_area = page.rect.width * page.rect.height
                                img_area = (r.x1 - r.x0) * (r.y1 - r.y0)
                                if img_area / page_area > ar_img_skip_thresh:
                                    continue  # Skip large decorative frames / full-page backgrounds
                                    
                                # Also skip images smaller than 10pt (tiny icons, noise)
                                w_img = r.x1 - r.x0
                                h_img = r.y1 - r.y0
                                if w_img < 10 or h_img < 10:
                                    continue
                                    
                                base_image = page.parent.extract_image(xref)
                                image_bytes = base_image["image"]
                                image_ext = base_image["ext"]
                                
                                temp_img_name = f"temp_ext_{unique_prefix}_{page_idx}_{img_idx}.{image_ext}"
                                temp_img_path = os.path.join(os.path.dirname(output_path), temp_img_name)
                                with open(temp_img_path, "wb") as f:
                                    f.write(image_bytes)
                                    
                                visual_elements.append({
                                    'path': temp_img_path,
                                    'y_rel': ((r.y0 + r.y1) / 2.0) / page.rect.height,
                                    'x_rel': ((r.x0 + r.x1) / 2.0) / page.rect.width,
                                    'type': 'extracted',
                                    'area_ratio': img_area / page_area
                                })
                                img_idx += 1
                        except Exception as ee:
                            logger.error(f"Failed standard raw extract on page {page_idx+1}: {ee}")
                        
            visual_elements.sort(key=lambda x: x['y_rel'])
            return visual_elements
            
        doc_pdf = fitz.open(input_path)
        total_pages_count = len(doc_pdf)
        combined_html = ["<html><head><meta charset='utf-8'></head><body>"]
        
        unique_prefix = os.path.splitext(os.path.basename(output_path))[0]
        all_page_images = {}

        # ─── LANGUAGE AUTO-DETECTION ─────────────────────────────────────────
        # Determine whether this is an Arabic (RTL) or English (LTR) document
        # so we can bifurcate prompts, fonts, and XML markers accordingly.
        def _detect_is_arabic(pdf_doc, user_lang, file_path):
            """Returns True if the document should be treated as Arabic/RTL."""
            if user_lang == 'ar':
                return True
            if user_lang == 'en':
                return False
            # lang == 'auto' → sample text from first 3 pages
            arabic_chars = 0
            total_chars  = 0
            for idx in range(min(3, len(pdf_doc))):
                txt = pdf_doc[idx].get_text('text')
                for ch in txt:
                    cp = ord(ch)
                    if 0x0600 <= cp <= 0x06FF or 0xFB50 <= cp <= 0xFDFF or 0xFE70 <= cp <= 0xFEFF:
                        arabic_chars += 1
                    if not ch.isspace():
                        total_chars += 1
            if total_chars > 20:
                return (arabic_chars / total_chars) >= 0.10
            # Scanned/image-only PDF → fall back to filename keywords
            name_lower = os.path.basename(file_path).lower()
            en_keywords = ['en', 'eng', 'english', 'activity', 'worksheet', 'exam']
            if any(kw in name_lower for kw in en_keywords):
                return False
            return True  # Default to Arabic (receipts, etc.)

        is_arabic = _detect_is_arabic(doc_pdf, lang, input_path)
        logger.info(f"Language detection → is_arabic={is_arabic} (lang hint='{lang}')")
        # ─────────────────────────────────────────────────────────────────────
        
        # Determine if this is a book/document (multi-page or dense text) vs a receipt
        total_text_chars = 0
        pages_to_check = min(3, total_pages_count)
        for idx in range(pages_to_check):
            total_text_chars += len(doc_pdf[idx].get_text('text').strip())
        avg_chars_per_page = (total_text_chars / pages_to_check) if pages_to_check > 0 else 0
        is_book_or_doc = (total_pages_count > 2) or (avg_chars_per_page > 600)
        logger.info(f"Document type detection -> avg_chars_per_page={avg_chars_per_page:.1f}, is_book_or_doc={is_book_or_doc}")
        
        all_full_page_images = {}
        
        for i in range(total_pages_count):
            logger.info(f"Sending page {i+1}/{total_pages_count} to Gemini Vision API...")
            if progress_callback:
                try:
                    progress_callback(i + 1, total_pages_count)
                except Exception as pe:
                    logger.error(f"Error calling progress callback: {pe}")
            page = doc_pdf[i]
            
            # Advanced crop/extract of visual elements (barcodes, QR codes, icons)
            all_page_images[i] = extract_page_visual_elements(page, unique_prefix, i, output_path, is_book_or_doc)
            
            # For English documents: also save a full-page PNG (used as fallback or for complex pages)
            if not is_arabic:
                try:
                    # High-quality 200 DPI render for the full page
                    pix_full = page.get_pixmap(matrix=fitz.Matrix(200/72, 200/72))
                    full_page_name = f"temp_fullpage_{unique_prefix}_{i}.png"
                    full_page_path = os.path.join(os.path.dirname(output_path), full_page_name)
                    pix_full.save(full_page_path)
                    all_full_page_images[i] = full_page_path
                except Exception as fpe:
                    logger.warning(f"Could not save full-page PNG for page {i+1}: {fpe}")
            
            # Render page as a high-quality JPEG image at 220 DPI for maximum OCR accuracy
            pix = page.get_pixmap(matrix=fitz.Matrix(220/72, 220/72))
            img_data = pix.tobytes("jpg", 95)
            img_b64 = base64.b64encode(img_data).decode('utf-8')
            
            # Calculate page aspect ratio/dimensions for context
            page_width_pt = page.rect.width
            page_height_pt = page.rect.height
            
            if is_arabic:
                # Is this a long book (many pages)? Adjust instructions accordingly.
                is_book = len(doc_pdf) > 10
                book_extra = (
                    "\n15. DECORATIVE BORDERS - CRITICAL: This is a multi-page book. "
                    "Pages contain decorative Islamic borders/frames. "
                    "IGNORE the borders completely. Do NOT output <img> tags for them. "
                    "Do NOT output ANY placeholder text like 'Decorative element not reproducible'. "
                    "Extract ONLY the text content.\n"
                    "16. NO LAYOUT TABLES: NEVER use <table> or margins to simulate page margins or borders. "
                    "Output the raw text as simple <p> tags."
                ) if is_book else ""
                
                prompt = (
                    "You are an expert Arabic document layout reconstruction assistant. "
                    "Analyze the provided Arabic page image and reconstruct its EXACT text, tables, headers, and visual layout into clean HTML.\n"
                    "Follow these rules STRICTLY:\n"
                    "1. RTL & ALIGNMENT — CRITICAL: All text must use inline CSS for direction and alignment. ALWAYS use style='text-align: right; direction: rtl;' for ALL paragraphs and body text. NEVER use center alignment for body text. Only center main big titles.\n"
                    "2. FONT SIZES & SPACING — CRITICAL: You MUST use inline CSS `font-size: ...px` on EVERY line to match the visual size exactly. Very large book titles (Thuluth/Naskh) MUST use font-size: 40-48px. Large headings use 24-32px. Body uses 14-18px. If there are vertical spaces between paragraphs, add `margin-bottom: 20px` etc.\n"
                    "3. BOLD & EMPHASIS: Preserve all bold text using <strong> or `font-weight: bold;`. Preserve italic using <em> tags.\n"
                    "4. TABLES: Use <table> for ANY multi-column layout. In RTL tables, Cell 0 (first <td>) = visually RIGHTMOST, last <td> = visually LEFTMOST.\n"
                    "5. PRECISION — CRITICAL: Extract every single character with 100% accuracy including all tashkeel (harakat) diacritics. NEVER omit, skip, or summarize text. NEVER guess or auto-complete words, proper nouns, brand names, or titles based on context (e.g. do not turn 'التأصيل' into 'التفاصيل' or similar common words). Read each character pixel by pixel.\n"
                    "   BEWARE of visually similar Arabic letters — read each character pixel by pixel:\n"
                    "   - (ص) vs (ض) vs (ط) vs (ظ): very different meanings\n"
                    "   - (أ) vs (إ) vs (آ) vs (ا): hamza position matters\n"
                    "   - (ه) vs (ة) vs (ح) vs (ج) vs (خ): distinguish carefully\n"
                    "   - (ر) vs (ز), (و) vs (ر), (ف) vs (ق): check dots precisely\n"
                    "   - (س) vs (ش), (ع) vs (غ), (ن) vs (ب) vs (ت) vs (ث): dot count is critical\n"
                    "   - For proper nouns, brand names, or titles: read them character by character, never guess.\n"
                    "6. NO HTML WRAPPER: Return ONLY the inner HTML inside a single <div>. No <html>, <head>, or <body> tags.\n"
                    "7. INLINE CSS: Use inline CSS for colors, sizes, weights. Keep visual appearance as close to original as possible.\n"
                    "8. MULTI-COLUMN LAYOUTS: Side-by-side elements, dual headers, key-value grids MUST be <table style='border:none;width:100%;'>.\n"
                    "9. IMAGES: Represent logos/QR/barcodes as <img src='image_0.png'>, <img src='image_1.png'> etc. (0-indexed by vertical order).\n"
                    "10. RTL TABLE CELL ORDER: Cell 0 = RIGHTMOST column, last cell = LEFTMOST column.\n"
                    "11. NO WORD MERGES: Always add spaces or <br> between adjacent text blocks. Never merge words together.\n"
                    "12. NO MARKDOWN: Return raw HTML only. No ```html fences.\n"
                    "13. NO INVENTED HARAKAT — CRITICAL: Do NOT add any tashkeel (harakat/diacritics) such as fatha (َ), damma (ُ), kasra (ِ), shadda (ّ), sukun (ْ), tanwin, etc. UNLESS they are CLEARLY and VISIBLY printed in the original image. Most Arabic documents (receipts, invoices, letters) have NO harakat at all. If the original text has no harakat, output it without any harakat. NEVER invent or add vowel marks that are not in the image.\n"
                    "14. NO SPACES BEFORE PUNCTUATION: Write titles properly like 'أ. إبراهيم' NOT 'أ . إبراهيم'. Never add spaces before dots, commas, or colons."
                    + book_extra
                )
            else:
                prompt = (
                    "You are an expert educational document layout reconstruction assistant. "
                    "Analyze the provided LTR page image (textbook, activity book, worksheet, or educational document) "
                    "and reconstruct its EXACT layout into clean, faithful HTML that matches the original as closely as possible.\n"
                    "Follow ALL of these rules STRICTLY — they are critical:\n\n"
                    "═══ RULE 1 — DIRECTION ═══\n"
                    "This is a LEFT-TO-RIGHT English document. Use style='direction: ltr; text-align: left;' on every container. NEVER use rtl anywhere.\n\n"
                    "═══ RULE 2 — FONT SIZES ═══\n"
                    "Match the visual font size hierarchy exactly:\n"
                    "  - Very large decorative titles (cover page book name): font-size: 36px+\n"
                    "  - Large headings (unit/lesson titles): font-size: 24-30px\n"
                    "  - Medium headings: font-size: 18-22px\n"
                    "  - Normal body text: font-size: 13-14px\n"
                    "  - Small captions/footnotes: font-size: 10-11px\n\n"
                    "═══ RULE 3 — COLORS & STYLING ═══\n"
                    "Preserve ALL text colors, background colors, bold (<strong>), and italic (<em>) exactly as in the original.\n\n"
                    "═══ RULE 4 — CRITICAL: ZERO WORD MERGING ═══\n"
                    "This is the MOST IMPORTANT rule. Text elements that are VISUALLY SEPARATE on the page MUST be in SEPARATE HTML elements. NEVER concatenate or merge them.\n"
                    "Concrete examples of FORBIDDEN merges:\n"
                    "  ✗ WRONG: 'JordanTEAMTogether' — ✓ CORRECT: <p>Jordan</p><p>TEAM Together</p>\n"
                    "  ✗ WRONG: 'Grade 3Semester 2' — ✓ CORRECT: <p>Grade 3</p><p>Semester 2</p>\n"
                    "  ✗ WRONG: 'three3' — ✓ CORRECT: <p>three 3</p>\n"
                    "  ✗ WRONG: 'ActivitiesW' — ✓ CORRECT: <p>Activities</p>\n"
                    "  ✗ WRONG: '2Look and write.' — ✓ CORRECT: <p>2 Look and write.</p>\n"
                    "Rule: Every visually distinct text block (even if spatially close) → its own <p> or <td>.\n\n"
                    "═══ RULE 5 — IMAGES ═══\n"
                    "Represent EVERY visible photo, illustration, drawing, picture, logo, or graphic as an <img> tag.\n"
                    "  - Number images 0, 1, 2, 3... ordered by their vertical position on the page (top to bottom, then left to right within same row).\n"
                    "  - Syntax: <img src='image_0.png'>, <img src='image_1.png'>, etc.\n"
                    "  - Place each <img> tag exactly where the image appears in the layout flow.\n"
                    "  - For image grids (e.g. 3 columns × 4 rows of pictures): use a <table> with <img> in each <td>.\n"
                    "  - For single inline images next to text: use <img> inline within the paragraph or table cell.\n"
                    "  - NEVER skip or omit any image.\n\n"
                    "═══ RULE 6 — TABLES FOR MULTI-COLUMN LAYOUTS ═══\n"
                    "Any content that appears side-by-side MUST use <table style='border:none; width:100%;'>:\n"
                    "  - Two-column author lists → 2-column borderless table\n"
                    "  - Page header with 3 sections (e.g. 'Welcome | My favourite day! | Lesson 1') → 3-column table\n"
                    "  - Image grids → table with images in cells\n"
                    "  - Numbered activity with image on one side, text on other → 2-column table\n\n"
                    "═══ RULE 7 — COVER / TITLE PAGES ═══\n"
                    "For cover pages with logo + title + grade + semester + author list:\n"
                    "  - Top logo/emblem image → <img src='image_0.png'>\n"
                    "  - Each title line in its own <p> or <h> tag (NEVER merge lines)\n"
                    "  - Author/committee names in a table or line-by-line with <br>\n"
                    "  - Publisher info block at bottom → preserved as-is\n\n"
                    "═══ RULE 8 — NUMBERED ACTIVITY BADGES ═══\n"
                    "When you see circled/badge numbers (①②③ or numbers in circles/boxes) before activity titles:\n"
                    "  - Render as: <p><strong>1 Look and write.</strong></p>\n"
                    "  - Always add a SPACE between the number and the activity text\n"
                    "  - NEVER merge number with text: '1Look' is WRONG, '1 Look' is CORRECT\n\n"
                    "═══ RULE 9 — FILL-IN-THE-BLANK LINES ═══\n"
                    "Preserve all blank answer lines exactly: '_______', '_ _ _ _', '.........' must appear verbatim. NEVER delete them.\n\n"
                    "═══ RULE 10 — IGNORE SHAPE METADATA ═══\n"
                    "Do NOT output descriptive names of empty shapes, borders, or decorative elements.\n"
                    "FORBIDDEN outputs: 'speech bubble pointer', 'callout', 'pointer', 'arrow shape', 'oval', 'rectangle'.\n"
                    "Only output ACTUAL printed text and image references.\n\n"
                    "═══ RULE 11 — PAGE NUMBERS ═══\n"
                    "Page numbers at the bottom → <p style='text-align:center;'>3</p>. Never merge with surrounding text.\n\n"
                    "═══ RULE 12 — OUTPUT FORMAT ═══\n"
                    "Return ONLY the inner HTML inside a single <div style='direction:ltr;'>. No <html>, <head>, <body> tags. No ```html fences."
                )
            
            payload = {
                "contents": [{
                    "parts": [
                        {"text": prompt},
                        {
                            "inlineData": {
                                "mimeType": "image/jpeg",
                                "data": img_b64
                            }
                        }
                    ]
                }]
            }
            models_to_try = ["gemini-2.5-flash", "gemini-2.0-flash", "gemini-1.5-flash", "gemini-flash-latest"]
            success = False
            content_text = ""
            
            for model_name in models_to_try:
                current_key = get_next_api_key(api_key) or api_key
                url = f"https://generativelanguage.googleapis.com/v1beta/models/{model_name}:generateContent?key={current_key}"
                max_retries = max(3, len(_api_keys_list) + 1)
                backoff_sec = 3
                model_failed = False
                
                for attempt in range(max_retries):
                    try:
                        response = requests.post(url, json=payload, timeout=180)
                        
                        # Handle API responses explicitly
                        if response.status_code != 200:
                            try:
                                err_data = response.json()
                                err_msg = err_data.get('error', {}).get('message', '')
                                err_status = err_data.get('error', {}).get('status', '')
                            except Exception:
                                err_msg = response.text
                                err_status = ''
                                
                            logger.warning(f"API error (status {response.status_code}) on {model_name}: {err_msg}")
                            
                            # 1. Permanent authentication or request format errors (400, 401, 403)
                            if response.status_code in (400, 401, 403) or "API key not valid" in err_msg:
                                logger.warning(f"Gemini API Key error (Auth/Invalid) on key ...{current_key[-4:] if current_key else ''}: {err_msg}")
                                if len(_api_keys_list) > 1:
                                    logger.info("Multiple keys available. Switching to next key due to Auth error...")
                                    current_key = get_next_api_key(api_key) or api_key
                                    url = f"https://generativelanguage.googleapis.com/v1beta/models/{model_name}:generateContent?key={current_key}"
                                    continue # Retry immediately with the new key!
                                else:
                                    raise RuntimeError(f"Gemini API Key is invalid or request is bad: {err_msg}")
                                
                            # 2. Rate limit or quota exhaustion (429)
                            if response.status_code == 429:
                                if err_status == "RESOURCE_EXHAUSTED" or "quota" in err_msg.lower():
                                    logger.warning(f"Model {model_name} quota exhausted on key ...{current_key[-4:] if current_key else ''}.")
                                    if len(_api_keys_list) > 1:
                                        logger.info("Multiple keys available. Switching to next key immediately...")
                                        current_key = get_next_api_key(api_key, _api_keys_list) or api_key
                                        url = f"https://generativelanguage.googleapis.com/v1beta/models/{model_name}:generateContent?key={current_key}"
                                        continue # Retry immediately with the new key!
                                    else:
                                        logger.warning(f"Only 1 key available. Waiting 60s...")
                                        if progress_callback:
                                            try: progress_callback(f"Waiting 60s for quota (page {i+1})...")
                                            except Exception: pass
                                        time.sleep(60)
                                        continue
                                else:
                                    # Other rate limits (e.g. temporary concurrency) - retry
                                    time.sleep(backoff_sec)
                                    backoff_sec = min(15, backoff_sec + 3)
                                    continue
                                    
                            # 3. Model Not Found (404) - skip to next model
                            if response.status_code == 404:
                                logger.warning(f"Model {model_name} not found (404). Trying next model...")
                                model_failed = True
                                break
                                
                            # 4. Service Unavailable / Overloaded (503) - skip to next model
                            if response.status_code == 503:
                                logger.warning(f"Model {model_name} is overloaded (503). Trying next model...")
                                model_failed = True
                                break
                                
                            # 5. Generic server errors - retry
                            if response.status_code >= 500:
                                time.sleep(backoff_sec)
                                backoff_sec = min(15, backoff_sec + 3)
                                continue
                                
                            response.raise_for_status()
                            
                        res_data = response.json()
                        candidates = res_data.get('candidates', [])
                        if candidates:
                            content_text = candidates[0].get('content', {}).get('parts', [{}])[0].get('text', '')
                            content_text = content_text.replace('```html', '').replace('```', '').strip()
                            success = True
                            break
                        else:
                            logger.warning(f"No response content for page {i+1} on model {model_name} attempt {attempt+1}")
                    except Exception as ex:
                        if "Gemini API Key is invalid" in str(ex):
                            raise ex
                        logger.warning(f"API attempt {attempt+1} failed for {model_name} on page {i+1}: {ex}")
                        if attempt == max_retries - 1:
                            break
                        time.sleep(backoff_sec)
                        backoff_sec = min(15, backoff_sec + 3)
                
                if model_failed:
                    continue  # Try next model in list
                if success:
                    break
            
            if success and content_text:
                if is_arabic:
                    # Replace non-breaking spaces with normal spaces
                    content_text = content_text.replace('&nbsp;', ' ')
                    # Remove spaces immediately following any HTML tag (fixes AI adding spaces to indent)
                    content_text = re.sub(r'(<[^>]+>)\s+', r'\1', content_text)
                    # Remove leading spaces at the very beginning of the string
                    content_text = re.sub(r'^\s+', '', content_text)
                    # Fix spaces before punctuation (e.g. 'قرن .' -> 'قرن.')
                    content_text = re.sub(r'\s+([\.،:,؛!؟])', r'\1', content_text)
                    # Fix flipped starting brackets often hallucinated by AI
                    content_text = re.sub(r'\](?=[\u0600-\u06FF])', '[', content_text)
                combined_html.append(f"<div class='page-break' data-page-index='{i}'>{content_text}</div>")
                if i < total_pages_count - 1:
                    logger.info("Sleeping for 4 seconds to respect API rate limits...")
                    time.sleep(4)
            else:
                raise RuntimeError(f"Failed to convert page {i+1} via Gemini API after trying all fallback models.")
        
        doc_pdf.close()
        combined_html.append("</body></html>")
        
        full_html = "\n".join(combined_html)
        
        # ─────────────────────────────────────────────
        # CUSTOM HIGH-FIDELITY HTML TO DOCX PARSER
        # ─────────────────────────────────────────────
        # ── Load pre-configured Arabic RTL template (industry standard approach) ──
        # Instead of creating a blank document and fighting XML ordering, we load a
        # template that Microsoft Word-compatible tools recognise as a native Arabic RTL
        # document. Every paragraph added to it inherits RTL automatically.
        _template_path = os.path.join(os.path.dirname(__file__), 'arabic_template.docx')
        if is_arabic and os.path.exists(_template_path):
            doc_word = Document(_template_path)
            # Remove any leftover paragraphs from the template
            for _p in doc_word.paragraphs:
                _p._element.getparent().remove(_p._element)
            logger.info("Loaded Arabic RTL template for document creation.")
        else:
            doc_word = Document()

        
        # Page size: A4 dimensions
        from docx.shared import Cm, Inches
        A4_WIDTH_CM  = 21.0
        A4_HEIGHT_CM = 29.7
        
        # Set default margins — tight to maximise content area
        MARGIN_PT = 24.0  # ~0.33 inch
        for section in doc_word.sections:
            section.page_width  = Cm(A4_WIDTH_CM)
            section.page_height = Cm(A4_HEIGHT_CM)
            section.top_margin    = Pt(MARGIN_PT)
            section.bottom_margin = Pt(MARGIN_PT)
            section.left_margin   = Pt(MARGIN_PT)
            section.right_margin  = Pt(MARGIN_PT)
            # For Arabic documents: set the entire section to Right-to-Left (BiDi)
            if is_arabic:
                sectPr = section._sectPr
                bidi_sect = OxmlElement('w:bidi')
                bidi_sect.set(qn('w:val'), '1')
                sectPr.append(bidi_sect)
            
        # Usable page width in points (for image/table sizing)
        PAGE_WIDTH_PT = A4_WIDTH_CM / 2.54 * 72 - 2 * MARGIN_PT  # ~491 pt
        
        # Set default font for document — Arabic gets Traditional Arabic 11.5pt, English gets Calibri 11pt
        style = doc_word.styles['Normal']
        font = style.font
        if is_arabic:
            font.name = 'Traditional Arabic'
            font.size = Pt(11.5)
            # ── CRITICAL: set w:bidi + w:jc start on the Normal STYLE itself ──
            # Without this, Word uses LTR for line-wrapping even if each paragraph
            # has w:bidi — the style-level direction wins for multi-line wrap.
            style_pPr = style.element.get_or_add_pPr()
            # Remove any existing bidi/jc on the style to avoid duplicates
            for old in style_pPr.findall(qn('w:bidi')):
                style_pPr.remove(old)
            for old in style_pPr.findall(qn('w:jc')):
                style_pPr.remove(old)
            style_bidi = OxmlElement('w:bidi')
            style_bidi.set(qn('w:val'), '1')
            style_pPr.append(style_bidi)
            style_jc = OxmlElement('w:jc')
            style_jc.set(qn('w:val'), 'start')
            style_pPr.append(style_jc)
            # Also set Arabic as the default run language in the Normal style
            style_rPr = style.element.get_or_add_rPr()
            for old in style_rPr.findall(qn('w:lang')):
                style_rPr.remove(old)
            lang_el = OxmlElement('w:lang')
            lang_el.set(qn('w:val'), 'ar-SA')
            lang_el.set(qn('w:bidi'), 'ar-SA')
            style_rPr.append(lang_el)
        else:
            font.name = 'Calibri'
            font.size = Pt(11)
        
        # Tight paragraph spacing to prevent vertical overflow
        style.paragraph_format.space_after = Pt(2)
        style.paragraph_format.space_before = Pt(0)
        style.paragraph_format.line_spacing = 1.15
        
        soup = BeautifulSoup(full_html, 'html.parser')
        
        def set_cell_background(cell, fill_color):
            tcPr = cell._tc.get_or_add_tcPr()
            shading = parse_xml(f'<w:shd {nsdecls("w")} w:fill="{fill_color}"/>')
            tcPr.append(shading)

        def set_cell_margins(cell, top=80, bottom=80, left=120, right=120):
            tcPr = cell._tc.get_or_add_tcPr()
            tcMar = OxmlElement('w:tcMar')
            for margin_name, val in [('top', top), ('bottom', bottom), ('left', left), ('right', right)]:
                node = OxmlElement(f'w:{margin_name}')
                node.set(qn('w:w'), str(val))
                node.set(qn('w:type'), 'dxa')
                tcMar.append(node)
            tcPr.append(tcMar)

        def set_cell_borders(cell, color="CCCCCC", sz="4", val="single"):
            tcPr = cell._tc.get_or_add_tcPr()
            tcBorders = OxmlElement('w:tcBorders')
            for border_name in ['top', 'left', 'bottom', 'right']:
                node = OxmlElement(f'w:{border_name}')
                node.set(qn('w:val'), val)
                node.set(qn('w:sz'), sz)
                node.set(qn('w:space'), '0')
                node.set(qn('w:color'), color)
                tcBorders.append(node)
            tcPr.append(tcBorders)

        def parse_css_color(style_str, property_name):
            if not style_str:
                return None
            pattern = rf'{property_name}\s*:\s*([^;]+)'
            match = re.search(pattern, style_str, re.IGNORECASE)
            if not match:
                return None
            val = match.group(1).strip()
            hex6 = re.match(r'^#([0-9a-fA-F]{6})$', val)
            if hex6:
                return hex6.group(1).upper()
            hex3 = re.match(r'^#([0-9a-fA-F]{3})$', val)
            if hex3:
                h = hex3.group(1)
                return f"{h[0]}{h[0]}{h[1]}{h[1]}{h[2]}{h[2]}".upper()
            rgb_match = re.match(r'^rgba?\(\s*(\d+)\s*,\s*(\d+)\s*,\s*(\d+)', val, re.IGNORECASE)
            if rgb_match:
                r = int(rgb_match.group(1))
                g = int(rgb_match.group(2))
                b = int(rgb_match.group(3))
                return f"{r:02X}{g:02X}{b:02X}"
            color_map = {
                'white': 'FFFFFF', 'black': '000000', 'gray': '808080', 'grey': '808080',
                'red': 'FF0000', 'green': '008000', 'blue': '0000FF', 'yellow': 'FFFF00'
            }
            if val.lower() in color_map:
                return color_map[val.lower()]
            return None

        def is_dark_color(hex_str):
            if not hex_str or len(hex_str) != 6:
                return False
            try:
                r = int(hex_str[0:2], 16)
                g = int(hex_str[2:4], 16)
                b = int(hex_str[4:6], 16)
                luminance = 0.2126 * r + 0.7152 * g + 0.0722 * b
                return luminance < 140
            except:
                return False

        current_page_index = 0
        page_imgs_used = set()

        def parse_inline_elements(element, paragraph, cell_bg=None, is_root=True):
            nonlocal page_imgs_used, current_page_index
            is_rtl = False
            is_center = False
            
            curr = element
            while curr and curr.name != '[document]':
                style_attr = curr.get('style', '')
                if 'direction: rtl' in style_attr or 'rtl' in style_attr:
                    is_rtl = True
                if 'text-align: center' in style_attr or 'center' in style_attr:
                    is_center = True
                elif 'text-align: right' in style_attr or 'right' in style_attr:
                    is_rtl = True
                if curr.get('align') == 'center' or curr.name == 'center':
                    is_center = True
                curr = curr.parent
                
            text_content = element.get_text(strip=True)
            
            # ARABIC ALIGNMENT OVERRIDE — CRITICAL:
            # Gemini often hallucinates text-align: center for body text and lists.
            # For Arabic documents, we ALWAYS force RIGHT alignment on everything EXCEPT
            # genuinely short centered titles (≤ 30 chars with no list markers).
            if is_arabic and is_center:
                # Only allow center for very short text that looks like a standalone title/number
                is_true_center = (len(text_content) <= 30 and
                                  not re.search(r'^[\s\u200e\u200f\u202a-\u202e]*[\d\u0660-\u0669][\s\-\.\)]+', text_content))
                if not is_true_center:
                    is_center = False  # Override: body text / lists → RIGHT

            if is_root:
                if is_arabic:
                    # Arabic documents: ALWAYS right-align EVERYTHING unless it's a true center title
                    if is_center:
                        paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
                    else:
                        paragraph.alignment = WD_ALIGN_PARAGRAPH.RIGHT
                elif is_center:
                    paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
                elif is_rtl:
                    paragraph.alignment = WD_ALIGN_PARAGRAPH.RIGHT
                else:
                    paragraph.alignment = WD_ALIGN_PARAGRAPH.LEFT
                    
                if is_arabic or is_rtl:
                    # ── CLEAN RTL pPr INJECTION ──────────────────────────────────────
                    # We REPLACE the entire pPr with a pre-built XML where order is
                    # HARDCODED correctly: bidi → spacing → jc
                    # This bypasses python-docx API ordering bugs where bidi ends up
                    # after spacing, causing Word to ignore it and wrap lines from left.
                    p_xml = paragraph._p
                    # Remove old pPr
                    old_pPr = p_xml.find(qn('w:pPr'))
                    if old_pPr is not None:
                        p_xml.remove(old_pPr)
                    # Build new pPr string with correct OOXML order
                    jc_val = 'center' if is_center else 'start'
                    new_pPr_xml = (
                        f'<w:pPr {nsdecls("w")}>'
                        f'<w:bidi w:val="1"/>'
                        f'<w:spacing w:after="40" w:before="0" w:line="276" w:lineRule="auto"/>'
                        f'<w:jc w:val="{jc_val}"/>'
                        f'</w:pPr>'
                    )
                    new_pPr = parse_xml(new_pPr_xml)
                    p_xml.insert(0, new_pPr)

                
            last_was_break = False
            for child in element.children:
                if child.name is None:
                    text = str(child)
                    if text.strip('\n'):
                        text = re.sub(r'\s+', ' ', text)
                        if not text.strip() and not paragraph.runs:
                            continue
                        last_was_break = False
                        curr_parent = child.parent
                        is_bold = False
                        is_italic = False
                        color = None
                        font_size = None
                        font_family = None
                        while curr_parent and curr_parent.name != '[document]':
                            if curr_parent.name in ['strong', 'b']:
                                is_bold = True
                            if curr_parent.name in ['em', 'i']:
                                is_italic = True
                            parent_style = curr_parent.get('style', '')
                            if 'font-weight: bold' in parent_style or 'bold' in parent_style:
                                is_bold = True
                            
                            parsed_color = parse_css_color(parent_style, 'color')
                            if parsed_color:
                                color = parsed_color
                                
                            if 'font-size:' in parent_style:
                                size_match = re.search(r'font-size:\s*([0-9.]+)\s*(px|pt|rem|em)?', parent_style)
                                if size_match:
                                    val = float(size_match.group(1))
                                    unit = size_match.group(2) or 'px'
                                    if unit == 'px':
                                        font_size = Pt(val * 0.75)
                                    elif unit == 'pt':
                                        font_size = Pt(val)
                                    elif unit in ('em', 'rem'):
                                        font_size = Pt(val * 12)
                            if 'font-family:' in parent_style:
                                fam_match = re.search(r'font-family:\s*([^;]+)', parent_style)
                                if fam_match:
                                    font_family = fam_match.group(1).strip().replace("'", "").replace('"', '').split(',')[0]
                            curr_parent = curr_parent.parent
                        
                        # Tighter metadata list rendering: insert break before bold receipt keys/colons if paragraph has runs
                        if is_bold and paragraph.runs:
                            last_run_text = paragraph.runs[-1].text if paragraph.runs else ""
                            if last_run_text.strip() and (':' in text or '：' in text or any(k in text for k in ["المركز", "رقم", "التاريخ", "الوقت", "الدافع", "الخدمة", "مكان", "قيمة", "حالة", "المجموع"])):
                                paragraph.add_run().add_break()
                                
                        run = paragraph.add_run(text)
                        
                        # White Text Safety Guard:
                        # If a text color is white but there is no dark background to contrast it,
                        # force it back to black so it doesn't render invisibly as white-on-white.
                        if color == 'FFFFFF':
                            if not cell_bg or not is_dark_color(cell_bg):
                                color = '000000'
                                
                        if is_bold:
                            run.bold = True
                        if is_italic:
                            run.italic = True
                        if color:
                            run.font.color.rgb = RGBColor.from_string(color)
                        if font_size:
                            run.font.size = font_size
                        
                        if is_arabic:
                            family_to_use = font_family if font_family else 'Traditional Arabic'
                            run.font.name = family_to_use
                            rPr = run._r.get_or_add_rPr()
                            rtl_xml = OxmlElement('w:rtl')
                            rPr.append(rtl_xml)
                            
                            # CS Font Family
                            rFonts = rPr.find(qn('w:rFonts'))
                            if rFonts is None:
                                rFonts = OxmlElement('w:rFonts')
                                rPr.append(rFonts)
                            rFonts.set(qn('w:cs'), family_to_use)
                            rFonts.set(qn('w:ascii'), family_to_use)
                            rFonts.set(qn('w:hAnsi'), family_to_use)
                            
                            # CS Font Size
                            if font_size:
                                szCs = OxmlElement('w:szCs')
                                szCs.set(qn('w:val'), str(int(font_size.pt * 2)))
                                rPr.append(szCs)
                        else:
                            family_to_use = font_family if font_family else 'Calibri'
                            run.font.name = family_to_use
                            rPr = run._r.get_or_add_rPr()
                            rFonts = rPr.find(qn('w:rFonts'))
                            if rFonts is None:
                                rFonts = OxmlElement('w:rFonts')
                                rPr.append(rFonts)
                            rFonts.set(qn('w:ascii'), family_to_use)
                            rFonts.set(qn('w:hAnsi'), family_to_use)
                elif child.name == 'br':
                    if not last_was_break:
                        paragraph.add_run().add_break()
                        last_was_break = True
                elif child.name == 'img':
                    last_was_break = False
                    src = child.get('src', '')
                    img_idx_match = re.search(r'image_(\d+)', src)
                    if img_idx_match and current_page_index in all_page_images:
                        try:
                            idx = int(img_idx_match.group(1))
                            page_imgs = all_page_images[current_page_index]
                            if idx < len(page_imgs):
                                elem = page_imgs[idx]
                                page_imgs_used.add(idx)
                                
                                # Setup picture paragraph alignment
                                if elem['type'] in ('qr', 'barcode'):
                                    paragraph.alignment = WD_ALIGN_PARAGRAPH.LEFT
                                    width_pt = 100
                                    if elem['type'] == 'barcode':
                                        width_pt = 140
                                else:
                                    paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
                                    # Dynamic width: image was rendered at 220/72 DPI.
                                    # Scale to fit PAGE_WIDTH_PT, preserving aspect ratio.
                                    try:
                                        from PIL import Image as PILImg
                                        with PILImg.open(elem['path']) as _pi:
                                            _pw, _ph = _pi.size
                                        # Convert rendered pixels → points (220 DPI render)
                                        img_w_pt = _pw * 72.0 / 220.0
                                        img_h_pt = _ph * 72.0 / 220.0
                                        # Scale down proportionally if wider than page
                                        if img_w_pt > PAGE_WIDTH_PT:
                                            scale = PAGE_WIDTH_PT / img_w_pt
                                            img_w_pt = PAGE_WIDTH_PT
                                            img_h_pt *= scale
                                        width_pt = max(img_w_pt, 50)
                                    except Exception:
                                        width_pt = min(PAGE_WIDTH_PT, 300)
                                
                                run = paragraph.add_run()
                                run.add_picture(elem['path'], width=Pt(width_pt))
                        except Exception as ie:
                            logger.error(f"Failed to insert nested image inside cell/p: {ie}")
                elif child.name is not None:
                    last_was_break = False
                    if child.name in ['p', 'div'] and paragraph.runs:
                        paragraph.add_run().add_break()
                    parse_inline_elements(child, paragraph, cell_bg=cell_bg, is_root=False)

        def process_node(node):
            nonlocal current_page_index, page_imgs_used
            if node.name == 'div' and node.get('data-page-index') is not None:
                try:
                    current_page_index = int(node.get('data-page-index'))
                except Exception:
                    pass
                
                # Recursively process children of page container (Gemini's <img> tags handle header logos inline)
                # Track which elements from all_page_images were actually inserted via <img> tags
                page_imgs_used.clear()
                
                def collect_used_img_indices(n):
                    """Walk HTML to find which image indices Gemini already placed via <img> tags."""
                    if n.name == 'img':
                        src = n.get('src', '')
                        m = re.search(r'image_(\d+)', src)
                        if m:
                            page_imgs_used.add(int(m.group(1)))
                    for ch in n.children:
                        if ch.name:
                            collect_used_img_indices(ch)
                
                collect_used_img_indices(node)
                
                # ── For English image-heavy pages: detect if Gemini output is almost all images ──
                # If the page HTML has little real text (< 30 chars) and we have a full-page PNG,
                # insert the full page as a single Word image for perfect visual fidelity.
                use_full_page_image = False
                if not is_arabic and current_page_index in all_full_page_images:
                    # Count text vs image content
                    all_text = node.get_text(separator=' ', strip=True)
                    img_count_in_html = len(node.find_all('img'))
                    text_length = len(all_text)
                    # Use full-page image if page is nearly all pictures (like activity book cover)
                    if img_count_in_html >= 1 and text_length < 80:
                        use_full_page_image = True
                
                if use_full_page_image:
                    # Insert the full-page PNG as a single centered image
                    try:
                        fp_path = all_full_page_images[current_page_index]
                        p_img = doc_word.add_paragraph()
                        p_img.alignment = WD_ALIGN_PARAGRAPH.CENTER
                        p_img.paragraph_format.space_before = Pt(0)
                        p_img.paragraph_format.space_after = Pt(0)
                        from PIL import Image as PILImg
                        with PILImg.open(fp_path) as _fpi:
                            _fw, _fh = _fpi.size
                        # Full page rendered at 200 DPI → scale to fit page width
                        fp_w_pt = _fw * 72.0 / 200.0
                        fp_h_pt = _fh * 72.0 / 200.0
                        if fp_w_pt > PAGE_WIDTH_PT:
                            scale = PAGE_WIDTH_PT / fp_w_pt
                            fp_w_pt = PAGE_WIDTH_PT
                            fp_h_pt *= scale
                        run_img = p_img.add_run()
                        run_img.add_picture(fp_path, width=Pt(fp_w_pt))
                        logger.info(f"Page {current_page_index+1}: inserted as full-page image ({fp_w_pt:.0f}pt wide).")
                    except Exception as fpe_err:
                        logger.warning(f"Full-page image insert failed: {fpe_err}. Falling back to HTML parse.")
                        for child in node.children:
                            if child.name:
                                process_node(child)
                else:
                    for child in node.children:
                        if child.name:
                            process_node(child)

                # ── Add a real Word page break AFTER each page (except the last) ──
                if current_page_index < total_pages_count - 1:
                    try:
                        p_break = doc_word.add_paragraph()
                        p_break.paragraph_format.space_before = Pt(0)
                        p_break.paragraph_format.space_after = Pt(0)
                        run_break = p_break.add_run()
                        from docx.enum.text import WD_BREAK
                        run_break.add_break(WD_BREAK.PAGE)
                        # Also force bidi on the break paragraph for Arabic
                        if is_arabic:
                            pPr = p_break._p.get_or_add_pPr()
                            bidi_br = OxmlElement('w:bidi')
                            pPr.append(bidi_br)
                    except Exception as pb_err:
                        logger.warning(f"Page break insertion failed: {pb_err}")
                        
                # Auto-insert ONLY elements that Gemini did NOT place via <img> tags (e.g. watermarked QR codes)
                # For both languages, skip auto-inserting large images (area_ratio > 0.15) to avoid decorative borders or background images.
                # If it's English, or if it's an Arabic book/document (is_book_or_doc is True), ONLY auto-insert QR/barcode elements (skip decorative/non-functional elements).
                # If it's a short Arabic receipt, auto-insert all unused elements (logos, stamps) as a safety fallback.
                if current_page_index in all_page_images:
                    for img_i, elem in enumerate(all_page_images[current_page_index]):
                        if img_i not in page_imgs_used and elem['y_rel'] >= 0.25:
                            elem_type = elem.get('type', 'extracted')
                            area_ratio = elem.get('area_ratio', 0)
                            
                            # 1. Skip any large unused image to avoid borders/backgrounds
                            if area_ratio > 0.15:
                                logger.info(f"Skipping auto-insert for large unused image (area_ratio={area_ratio:.2f})")
                                continue
                                
                            # 2. Skip non-essential elements for English or Arabic books/documents
                            if not is_arabic or is_book_or_doc:
                                if elem_type not in ('qr', 'barcode'):
                                    continue

                            placed = False
                            
                            # Integrate QR codes directly inside the signature table (second column, which is left-aligned in RTL)
                            # This aligns it side-by-side with signature block and prevents breaking onto page 2!
                            if is_arabic and elem['type'] in ('qr', 'barcode') and doc_word.tables:
                                last_table = doc_word.tables[-1]
                                if len(last_table.columns) >= 2:
                                    cell = last_table.cell(0, 1)
                                    p = cell.add_paragraph()
                                    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                                    p.paragraph_format.space_before = Pt(4)
                                    p.paragraph_format.space_after = Pt(2)
                                    run = p.add_run()
                                    try:
                                        width_pt = 60  # Ultra-compact premium size
                                        if elem['type'] == 'barcode':
                                            width_pt = 100
                                        run.add_picture(elem['path'], width=Pt(width_pt))
                                        placed = True
                                        page_imgs_used.add(img_i)
                                    except Exception as ie:
                                        logger.error(f"Failed to insert QR into signature table: {ie}")
                                        
                            if not placed:
                                p = doc_word.add_paragraph()
                                p.paragraph_format.space_before = Pt(2)
                                p.paragraph_format.space_after = Pt(2)
                                if elem['type'] in ('qr', 'barcode'):
                                    p.alignment = WD_ALIGN_PARAGRAPH.LEFT
                                else:
                                    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                                run = p.add_run()
                                try:
                                    width_pt = 60
                                    if elem['type'] == 'barcode':
                                        width_pt = 100
                                    run.add_picture(elem['path'], width=Pt(width_pt))
                                    page_imgs_used.add(img_i)
                                except Exception as e:
                                    logger.error(f"Failed to auto-insert footer element: {e}")
                return
            
            if node.name in ['p', 'div'] and not node.find(['p', 'div', 'table', 'ul', 'ol']):
                text_content = node.get_text().strip()
                if text_content:
                    p = doc_word.add_paragraph()
                    
                    # Intercept diagonal watermarks extracted as text (e.g. "نسخة الدافع")
                    # Render them centered, bold, elegant light gray, and compact to avoid page overflow
                    is_watermark = any(term in text_content for term in ["نسخة الدافع", "نسخة للدافع", "نسخة العميل", "نسخة الحفظ", "نسخة المحاسب", "نسخة الأرشيف"])
                    is_metadata = any(term in text_content for term in ["تاريخ ووقت الطباعة", "اسم المستخدم"])
                    
                    if is_watermark:
                        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                        p.paragraph_format.space_before = Pt(4)
                        p.paragraph_format.space_after = Pt(4)
                        
                        run = p.add_run(text_content)
                        run.font.name = 'Traditional Arabic'
                        run.font.size = Pt(14)
                        run.bold = True
                        run.font.color.rgb = RGBColor.from_string("999999")
                        
                        rPr = run._r.get_or_add_rPr()
                        rtl_xml = OxmlElement('w:rtl')
                        rPr.append(rtl_xml)
                    elif is_metadata:
                        # Make print details extremely compact and elegant grey to prevent page breaks
                        p.alignment = WD_ALIGN_PARAGRAPH.LEFT
                        p.paragraph_format.space_before = Pt(1)
                        p.paragraph_format.space_after = Pt(1)
                        p.paragraph_format.line_spacing = 1.0
                        
                        run = p.add_run(text_content)
                        run.font.name = 'Traditional Arabic'
                        run.font.size = Pt(8.5)
                        run.font.color.rgb = RGBColor.from_string("777777")
                        
                        rPr = run._r.get_or_add_rPr()
                        rtl_xml = OxmlElement('w:rtl')
                        rPr.append(rtl_xml)
                    else:
                        parse_inline_elements(node, p)
            elif node.name in ['h1', 'h2', 'h3', 'h4', 'h5', 'h6']:
                p = doc_word.add_paragraph()
                parse_inline_elements(node, p)
                level_sizes = {'h1': 28, 'h2': 22, 'h3': 18, 'h4': 16}
                size = level_sizes.get(node.name, 14)
                for run in p.runs:
                    run.bold = True
                    if not run.font.size:
                        run.font.size = Pt(size)
                        if is_arabic:
                            rPr = run._r.get_or_add_rPr()
                            szCs = OxmlElement('w:szCs')
                            szCs.set(qn('w:val'), str(int(size * 2)))
                            rPr.append(szCs)
            elif node.name == 'table':
                table_style = node.get('style', '')
                has_border = True
                if 'border: none' in table_style or 'border: 0' in table_style or node.get('border') == '0':
                    has_border = False
                # Also check for border:none (no space)
                if re.search(r'border\s*:\s*none|border\s*:\s*0', table_style, re.IGNORECASE):
                    has_border = False
                
                rows = node.find_all('tr')
                if not rows:
                    return
                
                # ── Resolve actual column count (accounting for colspan) ──
                max_cols = 0
                for row in rows:
                    span_sum = 0
                    for cell_tag in row.find_all(['td', 'th']):
                        try:
                            span_sum += int(cell_tag.get('colspan', 1))
                        except Exception:
                            span_sum += 1
                    max_cols = max(max_cols, span_sum)
                    
                if max_cols == 0:
                    return

                # ── Parse percentage column widths from first row (for English tables) ──
                col_widths_pt = None
                if not is_arabic:
                    first_row_cells = rows[0].find_all(['td', 'th']) if rows else []
                    if first_row_cells and len(first_row_cells) == max_cols:
                        widths = []
                        all_have_width = True
                        for ctag in first_row_cells:
                            cst = ctag.get('style', '')
                            wm = re.search(r'width\s*:\s*([0-9.]+)\s*(%|px|pt)', cst, re.IGNORECASE)
                            if wm:
                                val = float(wm.group(1))
                                unit = wm.group(2)
                                if unit == '%':
                                    widths.append(val / 100.0 * PAGE_WIDTH_PT)
                                elif unit == 'px':
                                    widths.append(val * 0.75)  # px→pt
                                else:
                                    widths.append(val)
                            else:
                                all_have_width = False
                                break
                        if all_have_width and sum(widths) > 0:
                            col_widths_pt = widths
                
                table = doc_word.add_table(rows=len(rows), cols=max_cols)
                if is_arabic:
                    tblPr = table._tbl.tblPr
                    if tblPr is not None:
                        bidi = OxmlElement('w:bidiVisual')
                        tblPr.append(bidi)
                table.autofit = False
                
                # Apply column widths if parsed
                if col_widths_pt:
                    from docx.oxml.ns import qn as _qn
                    tbl = table._tbl
                    tblGrid = tbl.find(_qn('w:tblGrid'))
                    if tblGrid is None:
                        tblGrid = OxmlElement('w:tblGrid')
                        tbl.insert(0, tblGrid)
                    # Clear existing gridCol elements
                    for gc in tblGrid.findall(_qn('w:gridCol')):
                        tblGrid.remove(gc)
                    for w_pt in col_widths_pt:
                        gc = OxmlElement('w:gridCol')
                        gc.set(_qn('w:w'), str(int(w_pt * 20)))  # twips
                        tblGrid.append(gc)
                else:
                    # Equal column widths
                    col_w_pt = PAGE_WIDTH_PT / max_cols
                    from docx.oxml.ns import qn as _qn
                    tbl = table._tbl
                    tblGrid = tbl.find(_qn('w:tblGrid'))
                    if tblGrid is None:
                        tblGrid = OxmlElement('w:tblGrid')
                        tbl.insert(0, tblGrid)
                    for gc in tblGrid.findall(_qn('w:gridCol')):
                        tblGrid.remove(gc)
                    for _ in range(max_cols):
                        gc = OxmlElement('w:gridCol')
                        gc.set(_qn('w:w'), str(int(col_w_pt * 20)))
                        tblGrid.append(gc)

                # ── Track merged cells (colspan / rowspan) ──
                occupied = {}  # (r, c) → True means already occupied by a span
                
                for r_idx, row in enumerate(rows):
                    tr_style = row.get('style', '')
                    bg_color = parse_css_color(tr_style, 'background-color')
                    
                    cell_tags = row.find_all(['td', 'th'])
                    c_cursor = 0  # tracks actual column position
                    for col in cell_tags:
                        # Skip occupied (already spanned) columns
                        while (r_idx, c_cursor) in occupied:
                            c_cursor += 1
                        if c_cursor >= max_cols:
                            break
                        
                        try:
                            colspan = int(col.get('colspan', 1))
                        except Exception:
                            colspan = 1
                        try:
                            rowspan = int(col.get('rowspan', 1))
                        except Exception:
                            rowspan = 1
                        
                        colspan = max(1, min(colspan, max_cols - c_cursor))
                        rowspan = max(1, rowspan)
                        
                        # Merge cells if colspan or rowspan > 1
                        if colspan > 1 or rowspan > 1:
                            try:
                                end_row = min(r_idx + rowspan - 1, len(rows) - 1)
                                end_col = min(c_cursor + colspan - 1, max_cols - 1)
                                # Ensure rows exist in the table (add if needed)
                                while len(table.rows) <= end_row:
                                    table.add_row()
                                cell = table.cell(r_idx, c_cursor).merge(
                                    table.cell(end_row, end_col)
                                )
                            except Exception:
                                cell = table.cell(r_idx, c_cursor)
                        else:
                            cell = table.cell(r_idx, c_cursor)
                        
                        # Mark spanned positions as occupied
                        for dr in range(rowspan):
                            for dc in range(colspan):
                                if dr > 0 or dc > 0:
                                    occupied[(r_idx + dr, c_cursor + dc)] = True
                        
                        cell_style = col.get('style', '')
                        col_bg = parse_css_color(cell_style, 'background-color') or bg_color
                                
                        if col_bg:
                            set_cell_background(cell, col_bg)
                            
                        set_cell_margins(cell)
                        if has_border:
                            set_cell_borders(cell)
                        else:
                            set_cell_borders(cell, val="none", sz="0", color="auto")
                        
                        # Set cell width via tcW XML for proportional columns
                        if col_widths_pt and c_cursor < len(col_widths_pt):
                            try:
                                tcPr = cell._tc.get_or_add_tcPr()
                                tcW = OxmlElement('w:tcW')
                                cell_w = sum(col_widths_pt[c_cursor:c_cursor + colspan])
                                tcW.set(qn('w:w'), str(int(cell_w * 20)))
                                tcW.set(qn('w:type'), 'dxa')
                                tcPr.append(tcW)
                            except Exception:
                                pass
                        
                        p = cell.paragraphs[0]
                        parse_inline_elements(col, p, cell_bg=col_bg)
                        
                        # Apply compact styling and prevent split for signature/terms table
                        cell_text = col.get_text()
                        if any(term in cell_text for term in ["أمين الصندوق", "محاسب القبض", "وصلني مبلغ"]):
                            try:
                                trPr = table.rows[r_idx]._tr.get_or_add_trPr()
                                trPr.append(parse_xml(r'<w:cantSplit xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"/>'))
                            except Exception:
                                pass
                                
                            for p_cell in cell.paragraphs:
                                p_cell.paragraph_format.space_before = Pt(1)
                                p_cell.paragraph_format.space_after = Pt(1)
                                p_cell.paragraph_format.line_spacing = 1.0
                                for run in p_cell.runs:
                                    run.font.size = Pt(9.5)
                        
                        c_cursor += colspan
            elif node.name in ['ul', 'ol']:
                idx = 1
                for li in node.find_all('li', recursive=False):
                    p = doc_word.add_paragraph()
                    bullet_text = "• " if node.name == 'ul' else f"{idx}. "
                    
                    parse_inline_elements(li, p)
                    
                    if p.runs:
                        p.runs[0].text = bullet_text + p.runs[0].text
                    else:
                        p.add_run(bullet_text)
                        
                    if is_arabic:
                        # For Arabic RTL: NO indent — let text flow flush to the right margin.
                        # right_indent pushes text away from the right edge which is unwanted.
                        pPr_li = p._p.get_or_add_pPr()
                        for old_bidi in pPr_li.findall(qn('w:bidi')):
                            pPr_li.remove(old_bidi)
                        bidi_li = OxmlElement('w:bidi')
                        pPr_li.append(bidi_li)
                    else:
                        p.paragraph_format.left_indent = Pt(20)
                        p.paragraph_format.first_line_indent = Pt(-15)
                        
                    idx += 1
            elif node.name == 'img':
                src = node.get('src', '')
                img_idx_match = re.search(r'image_(\d+)', src)
                if img_idx_match and current_page_index in all_page_images:
                    try:
                        idx = int(img_idx_match.group(1))
                        page_imgs = all_page_images[current_page_index]
                        if idx < len(page_imgs):
                            elem = page_imgs[idx]
                            p = doc_word.add_paragraph()
                            # Logo/emblem = center, QR/barcode = left
                            if elem['type'] in ('qr', 'barcode'):
                                p.alignment = WD_ALIGN_PARAGRAPH.LEFT
                                width_pt = 100
                            else:
                                p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                                # Scale image to fit page width, preserve aspect ratio
                                try:
                                    from PIL import Image as PILImg
                                    with PILImg.open(elem['path']) as _pi:
                                        _pw, _ph = _pi.size
                                    img_w_pt = _pw * 72.0 / 220.0
                                    img_h_pt = _ph * 72.0 / 220.0
                                    # Scale down to fit within page width
                                    if img_w_pt > PAGE_WIDTH_PT:
                                        scale = PAGE_WIDTH_PT / img_w_pt
                                        img_w_pt = PAGE_WIDTH_PT
                                        img_h_pt *= scale
                                    width_pt = max(img_w_pt, 50)
                                except Exception:
                                    width_pt = min(PAGE_WIDTH_PT, 300)
                            run = p.add_run()
                            run.add_picture(elem['path'], width=Pt(width_pt))
                    except Exception as ie:
                        logger.error(f"Failed to insert barcode/logo image: {ie}")
            else:
                for child in node.children:
                    if child.name:
                        process_node(child)

        process_node(soup)
        
        # Save to the output path
        doc_word.save(output_path)
        
        # ── DEFINITIVE ARABIC RTL FIX: Direct ZIP-level XML patch ──────────────────
        # Why python-docx API fails: it appends w:bidi AFTER w:spacing in pPr,
        # but OOXML schema requires: bidi → spacing → jc (in that exact order).
        # Word silently ignores out-of-order w:bidi, causing lines to wrap from LEFT.
        # Solution: Open the DOCX as a ZIP and patch every pPr directly with lxml.
        if is_arabic:
            try:
                import zipfile as _zf
                import io as _io
                from lxml import etree as _ET

                _W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'

                def _qw(tag):
                    return f'{{{_W}}}{tag}'

                # OOXML schema: these tags must come AFTER w:bidi in pPr
                _AFTER_BIDI = {_qw(t) for t in [
                    'adjustRightInd', 'snapToGrid', 'spacing', 'ind',
                    'contextualSpacing', 'mirrorIndents', 'suppressOverlap',
                    'jc', 'textDirection', 'textAlignment', 'textboxTightWrap',
                    'outlineLvl', 'divId', 'cnfStyle', 'rPr', 'sectPr', 'pPrChange'
                ]}
                # OOXML schema: these tags must come AFTER w:jc in pPr
                _AFTER_JC = {_qw(t) for t in [
                    'textDirection', 'textAlignment', 'textboxTightWrap',
                    'outlineLvl', 'divId', 'cnfStyle', 'rPr', 'sectPr', 'pPrChange'
                ]}

                def _patch_pPr_base(pPr):
                    # Remove existing bidi and jc
                    for old in pPr.findall(_qw('bidi')):
                        pPr.remove(old)
                    for old in pPr.findall(_qw('jc')):
                        pPr.remove(old)
                    # Insert w:bidi at the correct position (before AFTER_BIDI tags)
                    bidi_el = _ET.Element(_qw('bidi'))
                    bidi_el.set(_qw('val'), '1')
                    pos_bidi = next((i for i, c in enumerate(pPr) if c.tag in _AFTER_BIDI), len(pPr))
                    pPr.insert(pos_bidi, bidi_el)
                    # Insert w:jc after w:bidi but before AFTER_JC tags
                    jc_el = _ET.Element(_qw('jc'))
                    jc_el.set(_qw('val'), 'start')
                    pos_jc = next((i for i, c in enumerate(pPr) if c.tag in _AFTER_JC), len(pPr))
                    pPr.insert(pos_jc, jc_el)

                def _patch_paragraph_pPr(pPr):
                    # Force custom ArabicBody style to prevent Word LTR Normal.dotm override
                    for old in pPr.findall(_qw('pStyle')):
                        pPr.remove(old)
                    pStyle = _ET.Element(_qw('pStyle'))
                    pStyle.set(_qw('val'), 'ArabicBody')
                    pPr.insert(0, pStyle)
                    _patch_pPr_base(pPr)

                def _patch_rPr(rPr):
                    if rPr.find(_qw('rtl')) is None:
                        _ET.SubElement(rPr, _qw('rtl'))
                    for old in rPr.findall(_qw('lang')):
                        rPr.remove(old)
                    lang_el = _ET.SubElement(rPr, _qw('lang'))
                    lang_el.set(_qw('val'), 'ar-SA')
                    lang_el.set(_qw('bidi'), 'ar-SA')

                buf = _io.BytesIO()
                with _zf.ZipFile(output_path, 'r') as zin, _zf.ZipFile(buf, 'w', _zf.ZIP_DEFLATED) as zout:
                    for item in zin.infolist():
                        data = zin.read(item.filename)

                        if item.filename == 'word/document.xml':
                            try:
                                tree = _ET.fromstring(data)
                                body = tree.find(f'{{{_W}}}body')
                                for p_el in body.findall(f'.//{_qw("p")}'):
                                    pPr = p_el.find(_qw('pPr'))
                                    if pPr is None:
                                        pPr = _ET.Element(_qw('pPr'))
                                        p_el.insert(0, pPr)
                                    _patch_paragraph_pPr(pPr)
                                    for r_el in p_el.findall(_qw('r')):
                                        rPr = r_el.find(_qw('rPr'))
                                        if rPr is None:
                                            rPr = _ET.Element(_qw('rPr'))
                                            r_el.insert(0, rPr)
                                        _patch_rPr(rPr)
                                sectPr = body.find(_qw('sectPr'))
                                if sectPr is not None:
                                    for old in sectPr.findall(_qw('bidi')):
                                        sectPr.remove(old)
                                    bidi_sect = _ET.SubElement(sectPr, _qw('bidi'))
                                    bidi_sect.set(_qw('val'), '1')
                                data = _ET.tostring(tree, xml_declaration=True, encoding='UTF-8', standalone=True)
                            except Exception as _e:
                                logger.warning(f'RTL patch document.xml: {_e}')

                        elif item.filename == 'word/styles.xml':
                            try:
                                tree = _ET.fromstring(data)
                                # Fix style-level pPr definitions (without modifying paragraph style pointers)
                                for pPr in tree.findall(f'.//{_qw("pPr")}'):
                                    _patch_pPr_base(pPr)
                                data = _ET.tostring(tree, xml_declaration=True, encoding='UTF-8', standalone=True)
                            except Exception as _e:
                                logger.warning(f'RTL patch styles.xml: {_e}')

                        zout.writestr(item, data)

                with open(output_path, 'wb') as _f:
                    _f.write(buf.getvalue())

                logger.info('Definitive Arabic RTL ZIP patch applied successfully.')
            except Exception as _rtl_err:
                logger.warning(f'Definitive Arabic RTL patch failed: {_rtl_err}')

        
        
        # Clean up temp extracted images safely
        for p_idx, imgs in all_page_images.items():
            for img_info in imgs:
                try:
                    img_path = img_info['path']
                    if os.path.exists(img_path):
                        os.remove(img_path)
                except Exception:
                    pass
        
        # Clean up full-page PNG temps
        for p_idx, fp_path in all_full_page_images.items():
            try:
                if os.path.exists(fp_path):
                    os.remove(fp_path)
            except Exception:
                pass
                    
        logger.info("Successfully converted PDF to DOCX using Premium BeautifulSoup Layout Reconstruction Engine!")
