# translations.py
import json
import os
import logging

# --- Global State ---
TRANSLATIONS = {} # Stores loaded translations, e.g., {"en": {"key": "value"}, "ar": {"key": "value"}}
CURRENT_LANGUAGE = "en" # Default language
DEFAULT_LANGUAGE = "en" # Define a default language
SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__)) # Define SCRIPT_DIR globally for easy access

# --- Load Translations ---
def load_language_file(lang_code):
    """Loads a specific language's translations from its JSON file."""
    global TRANSLATIONS
    filename = f"{lang_code}.json"
    file_path = os.path.join(SCRIPT_DIR, filename)

    if not os.path.exists(file_path):
        logging.error(f"Translation file not found: {file_path}")
        return False

    try:
        with open(file_path, 'r', encoding='utf-8') as f:
            TRANSLATIONS[lang_code] = json.load(f)
        logging.info(f"Translations for '{lang_code}' loaded successfully from {file_path}.")
        return True
    except json.JSONDecodeError:
        logging.error(f"Error decoding JSON from translation file '{filename}'.")
        TRANSLATIONS[lang_code] = {} # Set empty dict on error to prevent repeated load attempts for bad file
        return False
    except Exception as e:
        logging.error(f"An unexpected error occurred loading translations for '{lang_code}': {e}", exc_info=True)
        TRANSLATIONS[lang_code] = {}
        return False

# --- Set Language ---
def set_language(lang_code):
    """Sets the current language for translations. Loads the language if not already loaded."""
    global CURRENT_LANGUAGE
    logging.debug(f"set_language: Called with lang_code='{lang_code}'. Current: '{CURRENT_LANGUAGE}'. Loaded: {list(TRANSLATIONS.keys())}")

    if lang_code not in TRANSLATIONS:
        logging.info(f"Language '{lang_code}' not loaded. Attempting to load.")
        if not load_language_file(lang_code):
            logging.warning(f"Failed to load translations for '{lang_code}'. Language not changed.")
            # Optionally, fall back to default or current language if critical
            if DEFAULT_LANGUAGE not in TRANSLATIONS: # Ensure default is loaded if target fails
                if not load_language_file(DEFAULT_LANGUAGE): # Try to load default
                    logging.error(f"CRITICAL: Default language '{DEFAULT_LANGUAGE}' also failed to load.")
                    # No translations available, this is a critical state.
                    # Depending on app requirements, could raise an error or use hardcoded fallbacks.
                    return False # Indicate failure
            CURRENT_LANGUAGE = DEFAULT_LANGUAGE # Fallback to default if target lang load fails
            logging.warning(f"Fell back to default language: {DEFAULT_LANGUAGE}")
            return False # Indicate failure to set desired language

    CURRENT_LANGUAGE = lang_code
    logging.info(f"Application language changed to: {lang_code}.")
    return True

# --- Get Translation ---
def get_translation(key, lang=None):
    """Gets the translation for a key in the current or specified language."""
    global CURRENT_LANGUAGE, TRANSLATIONS, DEFAULT_LANGUAGE

    target_lang = lang if lang and lang in TRANSLATIONS else CURRENT_LANGUAGE
    logging.debug(f"get_translation: key='{key}', target_lang='{target_lang}'. Current global lang='{CURRENT_LANGUAGE}'.")

    # Ensure the target language is loaded, if not, try to load it.
    # This is a safeguard, set_language should handle primary loading.
    if target_lang not in TRANSLATIONS:
        logging.warning(f"get_translation: Target language '{target_lang}' not loaded. Attempting to load.")
        if not load_language_file(target_lang):
            logging.warning(f"get_translation: Failed to load '{target_lang}'. Falling back to default language '{DEFAULT_LANGUAGE}' for this lookup.")
            target_lang = DEFAULT_LANGUAGE # Fallback to default for this lookup
            if DEFAULT_LANGUAGE not in TRANSLATIONS: # Ensure default is loaded
                 if not load_language_file(DEFAULT_LANGUAGE):
                    logging.error(f"CRITICAL: Default language '{DEFAULT_LANGUAGE}' failed to load in get_translation.")
                    return f"_{key}_ERR_NO_LANG_" # Critical error, no translations

    # Attempt to get translation for the target language
    translation = TRANSLATIONS.get(target_lang, {}).get(key)
    if translation is not None:
        return translation

    # Fallback 1: Try Default language if target language failed and wasn't Default
    if target_lang != DEFAULT_LANGUAGE:
        logging.warning(f"Translation missing for key '{key}' in '{target_lang}'. Falling back to {DEFAULT_LANGUAGE}.")
        default_translation = TRANSLATIONS.get(DEFAULT_LANGUAGE, {}).get(key)
        if default_translation is not None:
            return default_translation

    # Fallback 2: Return the key itself marked if not found anywhere
    logging.warning(f"Translation missing for key='{key}' in '{target_lang}' and also in {DEFAULT_LANGUAGE} fallback.")
    return f"_{key}_"

# --- Initial Load ---
# Load the default language on startup.
if not load_language_file(DEFAULT_LANGUAGE):
    logging.critical(f"Failed to load default language '{DEFAULT_LANGUAGE}' on startup. Application may not function correctly.")
    # Handle this critical failure, e.g., by exiting or using a minimal set of hardcoded English strings.
    # For now, TRANSLATIONS[DEFAULT_LANGUAGE] might be an empty dict if load_language_file set it.
    if DEFAULT_LANGUAGE not in TRANSLATIONS: # If load_language_file didn't even create the key
        TRANSLATIONS[DEFAULT_LANGUAGE] = {} # Ensure the key exists to prevent crashes in get_translation

# Ensure CURRENT_LANGUAGE is set to something valid on startup,
# even if default language loading failed.
if DEFAULT_LANGUAGE in TRANSLATIONS:
    CURRENT_LANGUAGE = DEFAULT_LANGUAGE
else:
    # This case means default language file is missing or corrupt AND load_language_file failed to create an empty dict.
    # This is a very critical state. For robustness, set a hardcoded language and provide minimal functionality.
    logging.error("CRITICAL: No languages could be loaded. Setting current language to 'en' nominally.")
    CURRENT_LANGUAGE = "en" # Nominal setting
    if "en" not in TRANSLATIONS: # Ensure 'en' key exists in TRANSLATIONS
        TRANSLATIONS["en"] = {}