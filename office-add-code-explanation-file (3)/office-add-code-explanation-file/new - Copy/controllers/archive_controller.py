import os
import logging

class ArchiveController:
    """
    Handles dynamic discovery of archive folders based on a template structure and on-disk state.
    """
    def __init__(self, structure, archives_path):
        # structure: nested dict of headers → subheaders → sections → subsections
        self.structure = structure
        self.archives_path = archives_path
        self.folder_cache = {}

    def get_dynamic_folder_options(self, base_folder_path, template_options):
        """
        Returns a sorted list of folder names found both in the template and on disk.

        Args:
            base_folder_path (str): Path on disk to scan.
            template_options (dict|list|None): Template-defined names.

        Returns:
            list[str]: Combined and sorted unique folder names.
        """
        if base_folder_path in self.folder_cache:
            logging.debug(f"Cache hit for '{base_folder_path}'")
            return self.folder_cache[base_folder_path]

        logging.debug(f"Cache miss for '{base_folder_path}'. Scanning disk.")
        disk_folders = set()
        template_folders = set()

        if os.path.isdir(base_folder_path): # Keep this check for the base path itself
            try:
                for entry in os.scandir(base_folder_path):
                    if entry.is_dir() and not entry.name.startswith('.'):
                        disk_folders.add(entry.name)
            except OSError as e:
                logging.warning(f"Could not scan directory '{base_folder_path}': {e}")
        else:
            logging.debug(f"Base folder path does not exist: {base_folder_path}")

        if isinstance(template_options, dict):
            template_folders = set(template_options.keys())
        elif isinstance(template_options, list):
            template_folders = set(template_options)

        combined = sorted(template_folders.union(disk_folders))
        logging.debug(f"Combined options for '{base_folder_path}': {combined}. Caching result.")
        self.folder_cache[base_folder_path] = combined
        return combined

    def clear_cache(self, path_prefix=None):
        """
        Clears the folder cache.

        Args:
            path_prefix (str, optional): If provided, only cache entries where the path
                                         starts with this prefix will be cleared.
                                         If None, the entire cache is cleared.
        """
        if path_prefix is None:
            self.folder_cache.clear()
            logging.info("ArchiveController cache fully cleared.")
        else:
            # Iterate over a copy of keys if modifying the dict
            for cached_path in list(self.folder_cache.keys()):
                if cached_path.startswith(path_prefix):
                    del self.folder_cache[cached_path]
            logging.info(f"ArchiveController cache cleared for prefix: {path_prefix}")