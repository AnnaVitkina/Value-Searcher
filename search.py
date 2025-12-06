"""
File Search Utility - EXACT MATCH (Case Insensitive)
Searches for EXACT word matches inside Excel (.xlsx, .xls), Word (.docx), and CSV files.
Example: searching for "Anna" will find "anna", "Anna" but NOT "Joanna" or "annabel"
"""

import os
import re
import warnings
from pathlib import Path
from typing import List, Optional, Set
from dataclasses import dataclass, field
from datetime import datetime

# Suppress openpyxl warnings about missing default styles
warnings.filterwarnings('ignore', category=UserWarning, module='openpyxl')

# ============================================================
# HARDCODE YOUR BASE PATH HERE
# ============================================================
# For Google Colab with Google Drive:
BASE_PATH = "/content/drive/Shared drives"

# For local Windows:
# BASE_PATH = "C:/Users/avitkin/Documents"

# Leave empty to use full paths:
# BASE_PATH = ""
# ============================================================

# Required libraries for reading Office files
try:
    import openpyxl
except ImportError:
    openpyxl = None

try:
    import xlrd
except ImportError:
    xlrd = None

try:
    from docx import Document
except ImportError:
    Document = None


# Supported file extensions
SUPPORTED_EXTENSIONS = {'.xlsx', '.xls', '.docx', '.csv'}


def convert_windows_path(path: str) -> str:
    """
    Convert Windows path to Google Drive compatible path.

    """
    # Remove common Windows drive prefixes
    prefixes_to_remove = [
        r'G:\Shared drives\\',
        r'G:\Shared drives/',
        r'G:/Shared drives/',
        r'G:/Shared drives\\',
        'G:\\Shared drives\\',
        'G:/Shared drives/',
    ]
    
    result = path
    for prefix in prefixes_to_remove:
        if result.startswith(prefix):
            result = result[len(prefix):]
            break
    
    # Replace backslashes with forward slashes
    result = result.replace('\\', '/')
    
    # Remove leading/trailing slashes
    result = result.strip('/')
    
    return result


@dataclass
class FolderResult:
    """Represents search results for a final folder."""
    folder_name: str
    folder_path: str
    all_files: List[str]
    searched_file: Optional[str]
    searched_file_modified: Optional[str]
    search_found: bool
    search_details: Optional[str]


def check_dependencies():
    """Check if required libraries are installed."""
    missing = []
    if openpyxl is None:
        missing.append("openpyxl")
    if xlrd is None:
        missing.append("xlrd")
    if Document is None:
        missing.append("python-docx")
    
    if missing:
        print("Missing required libraries. Install them with:")
        print(f"  pip install {' '.join(missing)}")
        return False
    return True


def is_exact_match(cell_value: str, search_value: str) -> bool:
    """
    Check if search_value is an EXACT match in cell_value (case insensitive).
    
    Examples:
        is_exact_match("Anna", "anna") -> True
        is_exact_match("anna", "Anna") -> True  
        is_exact_match("Joanna", "anna") -> False
        is_exact_match("annabel", "anna") -> False
        is_exact_match("Hello Anna!", "anna") -> True
        is_exact_match("Anna,Bob", "anna") -> True
    """
    # Use word boundary regex for exact matching (case insensitive)
    # \b matches word boundaries (start/end of word)
    pattern = r'\b' + re.escape(search_value) + r'\b'
    return bool(re.search(pattern, cell_value, re.IGNORECASE))


def search_in_xlsx(file_path: Path, search_value: str) -> Optional[str]:
    """Search for EXACT value match inside an .xlsx file."""
    if openpyxl is None:
        return None
    
    try:
        workbook = openpyxl.load_workbook(file_path, read_only=True, data_only=True)
        matches = []
        
        for sheet_name in workbook.sheetnames:
            sheet = workbook[sheet_name]
            for row in sheet.iter_rows():
                for cell in row:
                    if cell.value is not None:
                        cell_str = str(cell.value)
                        if is_exact_match(cell_str, search_value):
                            matches.append(f"Sheet '{sheet_name}', Cell {cell.coordinate}")
        
        workbook.close()
        
        if matches:
            return "; ".join(matches[:5]) + (f" (+{len(matches)-5} more)" if len(matches) > 5 else "")
        return None
        
    except Exception:
        return None


def search_in_xls(file_path: Path, search_value: str) -> Optional[str]:
    """Search for EXACT value match inside an .xls file (older Excel format)."""
    if xlrd is None:
        return None
    
    try:
        workbook = xlrd.open_workbook(str(file_path))
        matches = []
        
        for sheet_idx in range(workbook.nsheets):
            sheet = workbook.sheet_by_index(sheet_idx)
            sheet_name = sheet.name
            
            for row_idx in range(sheet.nrows):
                for col_idx in range(sheet.ncols):
                    cell_value = sheet.cell_value(row_idx, col_idx)
                    if cell_value:
                        cell_str = str(cell_value)
                        if is_exact_match(cell_str, search_value):
                            col_letter = xlrd.colname(col_idx)
                            matches.append(f"Sheet '{sheet_name}', Cell {col_letter}{row_idx + 1}")
        
        if matches:
            return "; ".join(matches[:5]) + (f" (+{len(matches)-5} more)" if len(matches) > 5 else "")
        return None
        
    except Exception:
        return None


def search_in_docx(file_path: Path, search_value: str) -> Optional[str]:
    """Search for EXACT value match inside a .docx file."""
    if Document is None:
        return None
    
    try:
        doc = Document(str(file_path))
        matches = []
        
        # Search in paragraphs
        for i, para in enumerate(doc.paragraphs, 1):
            text = para.text
            if is_exact_match(text, search_value):
                preview = text[:50] + "..." if len(text) > 50 else text
                matches.append(f"Paragraph {i}: '{preview}'")
        
        # Search in tables
        for table_idx, table in enumerate(doc.tables, 1):
            for row_idx, row in enumerate(table.rows):
                for cell_idx, cell in enumerate(row.cells):
                    text = cell.text
                    if is_exact_match(text, search_value):
                        matches.append(f"Table {table_idx}, Row {row_idx + 1}, Cell {cell_idx + 1}")
        
        if matches:
            return "; ".join(matches[:5]) + (f" (+{len(matches)-5} more)" if len(matches) > 5 else "")
        return None
        
    except Exception:
        return None


def search_in_csv(file_path: Path, search_value: str) -> Optional[str]:
    """Search for EXACT value match inside a .csv file."""
    
    try:
        # Read file as binary first, then decode
        file_path_str = str(file_path)
        
        # Read raw bytes
        with open(file_path_str, 'rb') as f:
            raw_data = f.read()
        
        # Try to decode with different encodings
        content = None
        for encoding in ['utf-8', 'utf-8-sig', 'latin-1', 'cp1252', 'iso-8859-1']:
            try:
                content = raw_data.decode(encoding)
                break
            except:
                continue
        
        if not content:
            # Last resort: decode with errors ignored
            content = raw_data.decode('utf-8', errors='ignore')
        
        if not content:
            return None
        
        matches = []
        
        # Search line by line (simpler and more reliable)
        lines = content.replace('\r\n', '\n').replace('\r', '\n').split('\n')
        
        for row_idx, line in enumerate(lines, 1):
            if not line.strip():
                continue
            
            # First check if the line might contain a match
            if is_exact_match(line, search_value):
                # Try to find which column
                cells = line.split(',')
                found_in_cell = False
                for col_idx, cell in enumerate(cells):
                    cell_clean = cell.strip().strip('"').strip("'")
                    if is_exact_match(cell_clean, search_value):
                        # Convert to Excel-like column letter
                        col_letter = ""
                        col_num = col_idx + 1
                        while col_num > 0:
                            col_num, remainder = divmod(col_num - 1, 26)
                            col_letter = chr(65 + remainder) + col_letter
                        matches.append(f"Row {row_idx}, Col {col_letter}")
                        found_in_cell = True
                        break
                
                if not found_in_cell:
                    # If we can't find specific column, just report the row
                    matches.append(f"Row {row_idx}")
        
        if matches:
            return "; ".join(matches[:5]) + (f" (+{len(matches)-5} more)" if len(matches) > 5 else "")
        return None
        
    except Exception as e:
        # Uncomment next line for debugging:
        # print(f"  [DEBUG] CSV error for {file_path}: {e}")
        return None


def search_in_file(file_path: Path, search_value: str) -> Optional[str]:
    """Search for EXACT value match in a file based on its extension."""
    extension = file_path.suffix.lower()
    
    if extension == '.xlsx':
        return search_in_xlsx(file_path, search_value)
    elif extension == '.xls':
        return search_in_xls(file_path, search_value)
    elif extension == '.docx':
        return search_in_docx(file_path, search_value)
    elif extension == '.csv':
        return search_in_csv(file_path, search_value)
    
    return None


def is_old_folder(folder: Path) -> bool:
    """Check if folder is named 'Old' (case-insensitive)."""
    return folder.name.lower() == 'old'


def has_non_old_subfolders(folder: Path) -> bool:
    """Check if folder has any subfolders (excluding 'Old' folders)."""
    try:
        for item in folder.iterdir():
            if item.is_dir() and not is_old_folder(item):
                return True
    except PermissionError:
        pass
    return False


def is_final_folder(folder: Path) -> bool:
    """Check if folder is a 'final' folder (has no subfolders except possibly 'Old')."""
    return not has_non_old_subfolders(folder)


def get_supported_files(folder: Path) -> List[Path]:
    """Get all supported files in a folder (not recursive)."""
    files = []
    try:
        for item in folder.iterdir():
            if item.is_file() and item.suffix.lower() in SUPPORTED_EXTENSIONS:
                files.append(item)
    except PermissionError:
        pass
    return files


def get_most_recent_file(files: List[Path]) -> Optional[Path]:
    """Get the most recently modified file from a list."""
    if not files:
        return None
    return max(files, key=lambda f: f.stat().st_mtime)


def find_all_final_folders(root_paths: List[str]) -> List[Path]:
    """Find all final folders (folders with no subfolders except 'Old') in the given paths."""
    final_folders: Set[Path] = set()
    
    for root_path in root_paths:
        root_path = root_path.strip()
        if not root_path:
            continue
            
        root = Path(root_path)
        
        if not root.exists():
            print(f"Warning: Folder not found: {root_path}")
            continue
        
        if not root.is_dir():
            print(f"Warning: Path is not a directory: {root_path}")
            continue
        
        print(f"  📂 Scanning: {root_path}")
        
        for folder in root.rglob('*'):
            if folder.is_dir():
                if is_old_folder(folder):
                    continue
                skip = False
                for parent in folder.parents:
                    if is_old_folder(parent):
                        skip = True
                        break
                if skip:
                    continue
                
                if is_final_folder(folder):
                    final_folders.add(folder)
        
        if is_final_folder(root):
            final_folders.add(root)
    
    return sorted(final_folders, key=lambda f: str(f).lower())


def search_in_final_folders(
    search_value: str,
    folder_paths: List[str]
) -> List[FolderResult]:
    """Search for EXACT value match in final folders."""
    
    results: List[FolderResult] = []
    final_folders = find_all_final_folders(folder_paths)
    
    if not final_folders:
        print("\n  ⚠️  No final folders found.")
        return results
    
    print(f"\n  📊 Found {len(final_folders)} final folder(s) to process...")
    print()
    
    for idx, folder in enumerate(final_folders, 1):
        progress = int((idx / len(final_folders)) * 20)
        bar = "█" * progress + "░" * (20 - progress)
        print(f"  ⏳ [{bar}] {idx}/{len(final_folders)} - {folder.name[:30]}", end='\r')
        
        files = get_supported_files(folder)
        all_file_names = sorted([f.name for f in files])
        
        if not files:
            results.append(FolderResult(
                folder_name=folder.name,
                folder_path=str(folder),
                all_files=[],
                searched_file=None,
                searched_file_modified=None,
                search_found=False,
                search_details=None
            ))
            continue
        
        most_recent = get_most_recent_file(files)
        
        if most_recent:
            mod_time = datetime.fromtimestamp(most_recent.stat().st_mtime)
            mod_time_str = mod_time.strftime("%Y-%m-%d %H:%M:%S")
            
            search_result = search_in_file(most_recent, search_value)
            
            results.append(FolderResult(
                folder_name=folder.name,
                folder_path=str(folder),
                all_files=all_file_names,
                searched_file=most_recent.name,
                searched_file_modified=mod_time_str,
                search_found=search_result is not None,
                search_details=search_result
            ))
        else:
            results.append(FolderResult(
                folder_name=folder.name,
                folder_path=str(folder),
                all_files=all_file_names,
                searched_file=None,
                searched_file_modified=None,
                search_found=False,
                search_details=None
            ))
    
    print()
    return results


def print_results(results: List[FolderResult], search_value: str) -> None:
    """Print search results in a formatted way."""
    if not results:
        print("\n❌ No folders processed.")
        return
    
    folders_with_matches = [r for r in results if r.search_found]
    
    print("\n")
    print("╔" + "═"*78 + "")
    print("║" + f"  🎯 EXACT MATCH RESULTS FOR: '{search_value}'".ljust(78) + "")
    print("╠" + "═"*78 + "")
    print("║" + f"  📁 Processed: {len(results)} final folder(s)".ljust(78) + "")
    print("║" + f"  ✅ Exact matches found in: {len(folders_with_matches)} folder(s)".ljust(78) + "")
    print("╚" + "═"*78 + "")
    
    if not folders_with_matches:
        print("\n  ⚠️  No exact matches found in any folder.")
        return
    
    print("\n")
    print("┌" + "─"*78 + "")
    print("│" + "  📂 FOLDERS WITH EXACT MATCHES".ljust(78) + "")
    print("└" + "─"*78 + "")
    
    for i, result in enumerate(folders_with_matches, 1):
        print(f"\n  ╭{'─'*74}")
        print(f"  │  {i}. 📁 {result.folder_name}".ljust(77) + "")
        print(f"  ├{'─'*74}")
        print(f"  │  📍 Path: {result.folder_path}".ljust(77) + "")
        print(f"  │".ljust(77) + "")
        print(f"  │  📋 Files in folder ({len(result.all_files)}):".ljust(77) + "")
        
        for fname in result.all_files:
            if fname == result.searched_file:
                line = f"  │      ★ {fname} ← SEARCHED (most recent)"
            else:
                line = f"  │      • {fname}"
            print(line.ljust(77) + "")
        
        print(f"  │".ljust(77))
        print(f"  ├{'─'*74}")
        
        details = result.search_details if result.search_details else ""
        print(f"  │  📄 Details: {details[:60]}{'...' if len(details) > 60 else ''}")
        print(f"  │".ljust(77))
        print(f"  ├{'─'*74}")
        
        full_file_path = f"{result.folder_path}/{result.searched_file}"
        print(f"  │  🎯 MATCH: {full_file_path}")
        print(f"  ╰{'─'*74}")


# Main function
if __name__ == "__main__":
    print("\n")
    print("  ╔═══════════════════════════════════════════════════════════════════")
    print("  ║                                                                   ")
    print("  ║   🎯  SEARCH UTILITY                                             ")
    print("  ║                                                                   ")
    print("  ║   📄 Searches in Excel, Word (.docx), and CSV files              ")
    print("  ║   📂 Shows all files, searches the most recent one                ")
    print("  ║   🚫 Skips subfolders named 'Old'                                 ")
    print("  ║                                                                   ")
    print("  ╚═══════════════════════════════════════════════════════════════════")
    print()
    
    if BASE_PATH:
        print(f"  📍 Base path: {BASE_PATH}")
        print("     (Enter folder names relative to this path)")
        print()
    
    if not check_dependencies():
        exit(1)
    
    search_value = ""
    while not search_value:
        search_value = input("  🎯 Enter the value to search for: ").strip()
        if not search_value:
            print("  ⚠️  Search value is required. Please try again.\n")
    
    folder_paths = []
    while not folder_paths:
        if BASE_PATH:
            print(f"\n  📁 Enter folder names (relative to base path)")
            print("     Press Enter on empty line when done:\n")
        else:
            print("\n  📁 Enter full folder paths to search")
            print("     Press Enter on empty line when done:\n")
        path_num = 1
        
        while True:
            folder = input(f"     [{path_num}] ➜ ").strip()
            if not folder:
                break
            
            # Convert Windows path if needed (e.g., G:\Shared drives\...)
            if folder.startswith("G:") or "\\" in folder:
                folder = convert_windows_path(folder)
                print(f"          → Converted to: {folder}")
            
            # Combine with BASE_PATH if set
            if BASE_PATH and not folder.startswith("/"):
                full_path = f"{BASE_PATH}/{folder}"
            else:
                full_path = folder
            folder_paths.append(full_path)
            path_num += 1
        
        if not folder_paths:
            print("  ⚠️  At least one folder path is required. Please try again.")
    
    print(f"\n  🚀 Searching for EXACT match of '{search_value}' in {len(folder_paths)} folder(s):")
    for fp in folder_paths:
        print(f"     • {fp}")
    print()
    
    try:
        results = search_in_final_folders(
            search_value=search_value,
            folder_paths=folder_paths
        )
        
        print_results(results, search_value)
        
        print("\n  ✅ Search completed!")
        print()
        
    except Exception as e:
        print(f"\n  ❌ An error occurred: {e}")


