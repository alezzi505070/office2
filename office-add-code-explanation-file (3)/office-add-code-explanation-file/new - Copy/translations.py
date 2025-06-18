# translations.py
import json
import os
import logging

# --- Global State ---
TRANSLATIONS = {}
CURRENT_LANGUAGE = "en" # Default language

# --- Load Translations ---
def load_translations(filename="translations.json"):
    """Loads translations from a JSON file."""
    global TRANSLATIONS, CURRENT_LANGUAGE
    try:
        # Determine path relative to this script file
        script_dir = os.path.dirname(os.path.abspath(__file__))
        file_path = os.path.join(script_dir, filename)

        if not os.path.exists(file_path):
            logging.error(f"Translation file not found at expected path: {file_path}")
            TRANSLATIONS = {"en": {}, "ar": {}} # Provide empty structure
            return

        with open(file_path, 'r', encoding='utf-8') as f:
            TRANSLATIONS = json.load(f)
        # Ensure both 'en' and 'ar' keys exist, even if empty
        if "en" not in TRANSLATIONS: TRANSLATIONS["en"] = {}
        if "ar" not in TRANSLATIONS: TRANSLATIONS["ar"] = {}
        logging.info(f"Translations loaded successfully from {file_path}. Current language: {CURRENT_LANGUAGE}")

    except FileNotFoundError:
        logging.error(f"Translation file '{filename}' not found.")
        TRANSLATIONS = {"en": {}, "ar": {}}
    except json.JSONDecodeError:
        logging.error(f"Error decoding JSON from translation file '{filename}'.")
        TRANSLATIONS = {"en": {}, "ar": {}}
    except Exception as e:
        logging.error(f"An unexpected error occurred loading translations: {e}", exc_info=True)
        TRANSLATIONS = {"en": {}, "ar": {}}

# --- Set Language ---
def set_language(lang_code):
    """Sets the current language for translations."""
    global CURRENT_LANGUAGE
    logging.debug(f"set_language: Called with lang_code='{lang_code}'. Current global CURRENT_LANGUAGE='{CURRENT_LANGUAGE}'. TRANSLATIONS keys: {list(TRANSLATIONS.keys()) if TRANSLATIONS else 'Not loaded'}")
    if lang_code in TRANSLATIONS:
        logging.debug(f"set_language: '{lang_code}' found in TRANSLATIONS. Preparing to set CURRENT_LANGUAGE.")
        CURRENT_LANGUAGE = lang_code
        logging.info(f"Application language changed to: {lang_code}. Global CURRENT_LANGUAGE is now '{CURRENT_LANGUAGE}'.")
        logging.debug("set_language: Returning True.")
        return True
    else:
        logging.warning(f"set_language: Attempted to set unsupported language: '{lang_code}'. TRANSLATIONS keys: {list(TRANSLATIONS.keys()) if TRANSLATIONS else 'Not loaded'}")
        logging.debug("set_language: Returning False.")
        return False

# --- Get Translation ---
def get_translation(key, lang=None):
    """Gets the translation for a key in the current or specified language."""
    global CURRENT_LANGUAGE, TRANSLATIONS
    logging.debug(f"get_translation: Requested key='{key}', lang_arg='{lang}'. Global CURRENT_LANGUAGE='{CURRENT_LANGUAGE}'.")
    # Use specified language if provided, otherwise use global current language
    target_lang = lang if lang and lang in TRANSLATIONS else CURRENT_LANGUAGE
    logging.debug(f"get_translation: Determined target_lang='{target_lang}' for key='{key}'.")

    try:
        # Attempt to get translation for the target language
        translation = TRANSLATIONS.get(target_lang, {}).get(key)
        if translation is not None:
            logging.debug(f"get_translation: Found primary translation for key='{key}' in target_lang='{target_lang}': '{translation}'.")
            return translation

        # Fallback 1: Try English if target language failed and wasn't English
        if target_lang != 'en':
            translation_en = TRANSLATIONS.get('en', {}).get(key)
            if translation_en is not None:
                logging.warning(f"Translation missing for key '{key}' in language '{target_lang}'. Falling back to English.")
                logging.debug(f"get_translation: Found English fallback for key='{key}': '{translation_en}'. Original target_lang was '{target_lang}'.")
                return translation_en

        # Fallback 2: Return the key itself marked if not found anywhere
        logging.warning(f"get_translation: Translation missing for key='{key}' in target_lang='{target_lang}' and also in English fallback. Global CURRENT_LANGUAGE='{CURRENT_LANGUAGE}'.")
        return f"_{key}_" # Mark missing keys

    except Exception as e:
        logging.error(f"get_translation: Error retrieving translation for key='{key}' in lang='{target_lang}'. Global CURRENT_LANGUAGE='{CURRENT_LANGUAGE}'. Error: {e}", exc_info=True)
        return f"_{key}_" # Return marked key on error

# --- Initial Load ---
load_translations()