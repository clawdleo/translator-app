"""
Translator Module - OPTIMIZED
-----------------------------
Fast translation with batch processing and googletrans primary.
Target: <5 minutes for 20MB/115 slides
"""

import time
import logging
from typing import Optional, List, Dict
from concurrent.futures import ThreadPoolExecutor, as_completed

logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

LANGUAGE_CODES = {
    'slovenian': 'sl',
    'croatian': 'hr',
    'serbian': 'sr',
    'english': 'en',
    'german': 'de',
    'italian': 'it',
    'french': 'fr',
    'spanish': 'es',
}


class Translator:
    """
    Fast translation engine with batching and caching.
    Primary: googletrans (fast, free)
    Fallback: DeepL API (reliable but slower)
    """
    
    def __init__(self, target_lang: str, deepl_api_key: Optional[str] = None, status_callback=None):
        self.target_lang = LANGUAGE_CODES.get(target_lang.lower(), target_lang.lower())
        self.deepl_api_key = deepl_api_key
        self.status_callback = status_callback
        self._cache: Dict[str, str] = {}
        self._gtrans = None
        self._use_googletrans = True
        self._init_googletrans()
        self._batch_queue: List[str] = []
        self._translation_count = 0
        
    def _update_status(self, msg):
        if self.status_callback:
            self.status_callback(msg)
    
    def _init_googletrans(self):
        """Initialize googletrans - much faster than DeepL."""
        try:
            from googletrans import Translator as GTranslator
            self._gtrans = GTranslator()
            # Test it
            test = self._gtrans.translate("hello", dest=self.target_lang)
            if test and test.text:
                logger.info("googletrans initialized successfully")
                self._use_googletrans = True
            else:
                raise Exception("googletrans test failed")
        except Exception as e:
            logger.warning(f"googletrans unavailable: {e}, using DeepL")
            self._use_googletrans = False
            self._gtrans = None
    
    def translate(self, text: str, max_retries: int = 2) -> str:
        """Translate single text with caching."""
        if not text or not text.strip():
            return text
        
        stripped = text.strip()
        if len(stripped) < 2:
            return text
        if stripped.replace(' ', '').replace('\n', '').replace('\t', '').isnumeric():
            return text
        
        # Check cache
        cache_key = text
        if cache_key in self._cache:
            return self._cache[cache_key]
        
        # Try googletrans first (much faster)
        if self._use_googletrans and self._gtrans:
            result = self._translate_googletrans(text, max_retries)
            if result and result != text:
                self._cache[cache_key] = result
                self._translation_count += 1
                return result
        
        # Fallback to DeepL
        if self.deepl_api_key:
            result = self._translate_deepl(text, max_retries)
            if result:
                self._cache[cache_key] = result
                self._translation_count += 1
                return result
        
        return text
    
    def translate_batch(self, texts: List[str]) -> List[str]:
        """
        Translate multiple texts efficiently.
        Uses batching for DeepL or parallel for googletrans.
        """
        if not texts:
            return texts
        
        results = [''] * len(texts)
        to_translate = []  # (index, text) pairs needing translation
        
        # Check cache first
        for i, text in enumerate(texts):
            if not text or not text.strip() or len(text.strip()) < 2:
                results[i] = text
            elif text in self._cache:
                results[i] = self._cache[text]
            else:
                to_translate.append((i, text))
        
        if not to_translate:
            return results
        
        self._update_status(f"Translating {len(to_translate)} text blocks...")
        
        if self._use_googletrans and self._gtrans:
            # Parallel translation with googletrans
            translated = self._batch_googletrans([t[1] for t in to_translate])
            for (idx, orig), trans in zip(to_translate, translated):
                results[idx] = trans
                self._cache[orig] = trans
        elif self.deepl_api_key:
            # Batch translation with DeepL (up to 50 at a time)
            translated = self._batch_deepl([t[1] for t in to_translate])
            for (idx, orig), trans in zip(to_translate, translated):
                results[idx] = trans
                self._cache[orig] = trans
        else:
            # No translator available, return originals
            for idx, text in to_translate:
                results[idx] = text
        
        self._translation_count += len(to_translate)
        return results
    
    def _translate_googletrans(self, text: str, max_retries: int) -> Optional[str]:
        """Fast translation using googletrans."""
        for attempt in range(max_retries):
            try:
                result = self._gtrans.translate(text, dest=self.target_lang)
                if result and result.text:
                    return result.text
            except Exception as e:
                logger.debug(f"googletrans attempt {attempt+1} failed: {e}")
                if attempt < max_retries - 1:
                    time.sleep(0.5)
                    # Reinitialize on failure
                    try:
                        from googletrans import Translator as GTranslator
                        self._gtrans = GTranslator()
                    except:
                        pass
        return None
    
    def _batch_googletrans(self, texts: List[str]) -> List[str]:
        """Parallel batch translation with googletrans."""
        results = [''] * len(texts)
        
        def translate_one(idx_text):
            idx, text = idx_text
            try:
                result = self._gtrans.translate(text, dest=self.target_lang)
                return idx, result.text if result and result.text else text
            except Exception as e:
                logger.debug(f"Batch googletrans failed for item {idx}: {e}")
                return idx, text
        
        # Use thread pool for parallel translation
        with ThreadPoolExecutor(max_workers=10) as executor:
            futures = [executor.submit(translate_one, (i, t)) for i, t in enumerate(texts)]
            for future in as_completed(futures):
                try:
                    idx, translated = future.result()
                    results[idx] = translated
                except Exception as e:
                    logger.error(f"Future failed: {e}")
        
        return results
    
    def _translate_deepl(self, text: str, max_retries: int) -> Optional[str]:
        """Single text translation with DeepL."""
        import httpx
        
        deepl_lang = self.target_lang.upper()
        
        for attempt in range(max_retries):
            try:
                response = httpx.post(
                    'https://api-free.deepl.com/v2/translate',
                    headers={
                        'Authorization': f'DeepL-Auth-Key {self.deepl_api_key}',
                        'Content-Type': 'application/json'
                    },
                    json={'text': [text], 'target_lang': deepl_lang},
                    timeout=30.0
                )
                
                if response.status_code == 200:
                    data = response.json()
                    if data.get('translations'):
                        return data['translations'][0]['text']
                elif response.status_code == 429:
                    time.sleep(2)
                    
            except Exception as e:
                logger.debug(f"DeepL attempt {attempt+1} failed: {e}")
                if attempt < max_retries - 1:
                    time.sleep(1)
        
        return None
    
    def _batch_deepl(self, texts: List[str], batch_size: int = 50) -> List[str]:
        """Batch translation with DeepL API (up to 50 texts per request)."""
        import httpx
        
        results = []
        deepl_lang = self.target_lang.upper()
        
        for i in range(0, len(texts), batch_size):
            batch = texts[i:i+batch_size]
            
            try:
                response = httpx.post(
                    'https://api-free.deepl.com/v2/translate',
                    headers={
                        'Authorization': f'DeepL-Auth-Key {self.deepl_api_key}',
                        'Content-Type': 'application/json'
                    },
                    json={'text': batch, 'target_lang': deepl_lang},
                    timeout=60.0
                )
                
                if response.status_code == 200:
                    data = response.json()
                    translations = data.get('translations', [])
                    for j, item in enumerate(translations):
                        results.append(item.get('text', batch[j]))
                    # Fill any missing
                    while len(results) < i + len(batch):
                        results.append(batch[len(results) - i])
                else:
                    logger.warning(f"DeepL batch failed: {response.status_code}")
                    results.extend(batch)  # Keep originals
                    
            except Exception as e:
                logger.error(f"DeepL batch error: {e}")
                results.extend(batch)  # Keep originals
            
            # Small delay between batches to avoid rate limiting
            if i + batch_size < len(texts):
                time.sleep(0.5)
        
        return results
    
    def get_stats(self) -> dict:
        return {
            'cached_translations': len(self._cache),
            'total_translated': self._translation_count,
            'target_language': self.target_lang,
            'engine': 'googletrans' if self._use_googletrans else 'deepl'
        }
