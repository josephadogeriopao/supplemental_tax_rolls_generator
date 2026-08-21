def get_asset_path(relative_path: str) -> str:
    """
    Resolves the absolute path for both development (PyCharm/CLI) and PyInstaller builds.
    Uses runtime entry point detection to ensure safe relative mapping.
    """
    # 1. Inside PyInstaller bundle: assets are directly in the temp root folder
    if hasattr(sys, "_MEIPASS"):
        return os.path.normpath(os.path.join(sys._MEIPASS, relative_path))

    # 2. Inside Local Dev: Detect where the main executing file (main.py/app.py) lives
    # This prevents subfolder nesting or current working directory from breaking paths
    entry_file = os.path.abspath(sys.argv[0] if sys.argv else __file__)
    entry_dir = os.path.dirname(entry_file)

    # Normalize folder array path to slice off 'src' if it's trapped in the path
    parts = entry_dir.split(os.sep)
    if "src" in parts:
        src_index = parts.index("src")
        project_root = os.sep.join(parts[:src_index])
        if project_root.endswith(":"):
            project_root += os.sep
    else:
        project_root = entry_dir

    return os.path.normpath(os.path.join(project_root, relative_path))
