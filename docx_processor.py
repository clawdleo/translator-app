"""
DOCX Processor Module
Handles text translation in Word documents while preserving formatting.
"""

import logging
from docx import Document
from docx.table import Table

logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)


class DOCXProcessor:
    def __init__(self, translator, status_callback=None):
        self.translator = translator
        self.status_callback = status_callback
        self.stats = {
            'paragraphs_processed': 0,
            'tables_processed': 0,
            'headers_processed': 0,
            'footers_processed': 0,
            'text_runs_translated': 0,
            'errors': []
        }
    
    def _update_status(self, message):
        if self.status_callback:
            self.status_callback(message)
        logger.info(message)
    
    def process_file(self, input_path: str, output_path: str) -> dict:
        try:
            self._update_status("Loading document...")
            doc = Document(input_path)
            
            # Process main body paragraphs
            total_paragraphs = len(doc.paragraphs)
            self._update_status(f"Processing {total_paragraphs} paragraphs...")
            
            for idx, paragraph in enumerate(doc.paragraphs):
                if idx % 10 == 0:
                    self._update_status(f"Processing paragraph {idx + 1}/{total_paragraphs}...")
                self._process_paragraph(paragraph)
                self.stats['paragraphs_processed'] += 1
            
            # Process tables
            if doc.tables:
                self._update_status(f"Processing {len(doc.tables)} tables...")
                for table in doc.tables:
                    self._process_table(table)
            
            # Process headers and footers
            self._update_status("Processing headers and footers...")
            for section in doc.sections:
                self._process_header_footer(section)
            
            # Save the document
            self._update_status("Saving translated document...")
            doc.save(output_path)
            
            self.stats['translator_stats'] = self.translator.get_stats()
            return self.stats
            
        except Exception as e:
            logger.error(f"Error processing file: {e}")
            self.stats['errors'].append(str(e))
            raise
    
    def _process_paragraph(self, paragraph) -> None:
        """Process a paragraph by translating its runs while preserving formatting."""
        runs = paragraph.runs
        if not runs:
            return
        
        original_texts = [run.text for run in runs]
        combined_text = ''.join(original_texts)
        
        if not combined_text.strip():
            return
        
        translated_text = self.translator.translate(combined_text)
        
        if translated_text == combined_text:
            return
        
        self._redistribute_text_to_runs(runs, original_texts, translated_text)
        self.stats['text_runs_translated'] += len(runs)
    
    def _redistribute_text_to_runs(self, runs, original_texts, translated_text) -> None:
        """Redistribute translated text back to runs proportionally."""
        total_original_len = sum(len(t) for t in original_texts)
        
        if total_original_len == 0:
            if runs:
                runs[0].text = translated_text
            return
        
        translated_len = len(translated_text)
        current_pos = 0
        
        for i, (run, orig_text) in enumerate(zip(runs, original_texts)):
            if i == len(runs) - 1:
                run.text = translated_text[current_pos:]
            else:
                proportion = len(orig_text) / total_original_len
                chars_for_run = int(translated_len * proportion)
                end_pos = current_pos + chars_for_run
                
                # Try to break at word boundary
                if end_pos < translated_len:
                    space_pos = translated_text.rfind(' ', current_pos, end_pos + 10)
                    if space_pos > current_pos:
                        end_pos = space_pos + 1
                
                run.text = translated_text[current_pos:end_pos]
                current_pos = end_pos
    
    def _process_table(self, table: Table) -> None:
        """Process all cells in a table."""
        try:
            for row in table.rows:
                for cell in row.cells:
                    for paragraph in cell.paragraphs:
                        self._process_paragraph(paragraph)
            self.stats['tables_processed'] += 1
        except Exception as e:
            error_msg = f"Error processing table: {e}"
            logger.warning(error_msg)
            self.stats['errors'].append(error_msg)
    
    def _process_header_footer(self, section) -> None:
        """Process headers and footers in a section."""
        try:
            # Process header
            if section.header:
                for paragraph in section.header.paragraphs:
                    self._process_paragraph(paragraph)
                self.stats['headers_processed'] += 1
            
            # Process footer
            if section.footer:
                for paragraph in section.footer.paragraphs:
                    self._process_paragraph(paragraph)
                self.stats['footers_processed'] += 1
                
        except Exception as e:
            logger.debug(f"Could not access header/footer: {e}")
