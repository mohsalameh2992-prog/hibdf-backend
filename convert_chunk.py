import sys
import logging
from spire.pdf import PdfDocument, FileFormat

# Set up simple logging
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)

def main():
    if len(sys.argv) < 3:
        logger.error("Usage: python convert_chunk.py <input_pdf> <output_docx>")
        sys.exit(1)
        
    input_pdf = sys.argv[1]
    output_docx = sys.argv[2]
    lang = sys.argv[3] if len(sys.argv) > 3 else "en"
    
    logger.info(f"Subprocess starting conversion: {input_pdf} -> {output_docx} (lang={lang})")
    
    def try_spire():
        spire_doc = PdfDocument()
        spire_doc.LoadFromFile(input_pdf)
        spire_doc.SaveToFile(output_docx, FileFormat.DOCX)
        spire_doc.Close()
        
    def try_pdf2docx():
        from pdf2docx import Converter
        cv = Converter(input_pdf)
        cv.convert(output_docx)
        cv.close()

    try:
        if lang == "ar":
            try:
                try_pdf2docx()
            except Exception as e:
                logger.warning(f"pdf2docx failed: {e}. Trying Spire fallback...")
                try_spire()
        else:
            try:
                try_spire()
            except Exception as e:
                logger.warning(f"Spire failed: {e}. Trying pdf2docx fallback...")
                try_pdf2docx()
                
        logger.info("Conversion completed successfully in subprocess")
        print("SUCCESS")
        sys.exit(0)
    except Exception as e:
        logger.error(f"Conversion failed in subprocess: {e}", exc_info=True)
        print(f"ERROR: {e}")
        sys.exit(1)

if __name__ == "__main__":
    main()
