"""
PPTX Processor Module
Handles surgical text translation in PowerPoint files.
"""

import logging
from pptx import Presentation
from pptx.shapes.group import GroupShape
from pptx.shapes.base import BaseShape
from pptx.table import Table
from pptx.text.text import TextFrame

logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)


class PPTXProcessor:
    def __init__(self, translator, status_callback=None):
        self.translator = translator
        self.status_callback = status_callback
        self.stats = {
            'slides_processed': 0,
            'shapes_processed': 0,
            'text_runs_translated': 0,
            'tables_processed': 0,
            'notes_translated': 0,
            'groups_traversed': 0,
            'errors': []
        }
    
    def _update_status(self, message):
        if self.status_callback:
            self.status_callback(message)
        logger.info(message)
    
    def process_file(self, input_path: str, output_path: str) -> dict:
        try:
            self._update_status(f"Loading presentation...")
            prs = Presentation(input_path)
            num_slides = len(prs.slides)
            self._update_status(f"Loaded {num_slides} slides")
            
            for slide_idx, slide in enumerate(prs.slides):
                try:
                    self._update_status(f"Processing slide {slide_idx + 1}/{num_slides}...")
                    self._process_slide(slide)
                    self.stats['slides_processed'] += 1
                except Exception as e:
                    error_msg = f"Error on slide {slide_idx + 1}: {e}"
                    logger.error(error_msg)
                    self.stats['errors'].append(error_msg)
            
            self._update_status(f"Saving translated presentation...")
            prs.save(output_path)
            
            self.stats['translator_stats'] = self.translator.get_stats()
            return self.stats
            
        except Exception as e:
            logger.error(f"Error processing file: {e}")
            self.stats['errors'].append(str(e))
            raise
    
    def _process_slide(self, slide) -> None:
        for shape in slide.shapes:
            self._process_shape(shape)
        self._process_notes(slide)
    
    def _process_shape(self, shape: BaseShape) -> None:
        try:
            if isinstance(shape, GroupShape):
                self.stats['groups_traversed'] += 1
                for child_shape in shape.shapes:
                    self._process_shape(child_shape)
                return
            
            if shape.has_table:
                self._process_table(shape.table)
                return
            
            if shape.has_text_frame:
                self._process_text_frame(shape.text_frame)
            
            self.stats['shapes_processed'] += 1
            
        except Exception as e:
            error_msg = f"Error processing shape: {e}"
            logger.warning(error_msg)
            self.stats['errors'].append(error_msg)
    
    def _process_text_frame(self, text_frame: TextFrame) -> None:
        for paragraph in text_frame.paragraphs:
            self._process_paragraph(paragraph)
    
    def _process_paragraph(self, paragraph) -> None:
        runs = list(paragraph.runs)
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
                
                if end_pos < translated_len:
                    space_pos = translated_text.rfind(' ', current_pos, end_pos + 10)
                    if space_pos > current_pos:
                        end_pos = space_pos + 1
                
                run.text = translated_text[current_pos:end_pos]
                current_pos = end_pos
    
    def _process_table(self, table: Table) -> None:
        try:
            for row in table.rows:
                for cell in row.cells:
                    if cell.text_frame:
                        self._process_text_frame(cell.text_frame)
            self.stats['tables_processed'] += 1
        except Exception as e:
            error_msg = f"Error processing table: {e}"
            logger.warning(error_msg)
            self.stats['errors'].append(error_msg)
    
    def _process_notes(self, slide) -> None:
        try:
            if not slide.has_notes_slide:
                return
            notes_slide = slide.notes_slide
            if notes_slide and notes_slide.notes_text_frame:
                self._process_text_frame(notes_slide.notes_text_frame)
                self.stats['notes_translated'] += 1
        except Exception as e:
            logger.debug(f"Could not access notes: {e}")
