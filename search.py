"""
File Search Utility
Searches for values inside Excel (.xlsx, .xls) and Word (.docx) files.
For each "final" folder (no subfolders), shows all files and searches only the most recently modified one.
"""

import os
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
BASE_PATH = "/content/drive/MyDrive"

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
SUPPORTED_EXTENSIONS = {'.xlsx', '.xls', '.docx'}


@dataclass
class FolderResult:
    """Represents search results for a final folder."""
    folder_name: str
    folder_path: str
    all_files: List[str]  # All file names in this folder
    searched_file: Optional[str]  # The most recent file that was searched
    searched_file_modified: Optional[str]  # Last modified date of searched file
    search_found: bool  # Whether search term was found
    search_details: Optional[str]  # Where the match was found


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


def search_in_xlsx(file_path: Path, search_value: str, case_sensitive: bool = False) -> Optional[str]:
    """Search for a value inside an .xlsx file."""
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
                        compare_value = cell_str if case_sensitive else cell_str.lower()
                        search_term = search_value if case_sensitive else search_value.lower()
                        
                        if search_term in compare_value:
                            matches.append(f"Sheet '{sheet_name}', Cell {cell.coordinate}")
        
        workbook.close()
        
        if matches:
            return "; ".join(matches[:5]) + (f" (+{len(matches)-5} more)" if len(matches) > 5 else "")
        return None
        
    except Exception:
        return None


def search_in_xls(file_path: Path, search_value: str, case_sensitive: bool = False) -> Optional[str]:
    """Search for a value inside an .xls file (older Excel format)."""
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
                        compare_value = cell_str if case_sensitive else cell_str.lower()
                        search_term = search_value if case_sensitive else search_value.lower()
                        
                        if search_term in compare_value:
                            col_letter = xlrd.colname(col_idx)
                            matches.append(f"Sheet '{sheet_name}', Cell {col_letter}{row_idx + 1}")
        
        if matches:
            return "; ".join(matches[:5]) + (f" (+{len(matches)-5} more)" if len(matches) > 5 else "")
        return None
        
    except Exception:
        return None


def search_in_docx(file_path: Path, search_value: str, case_sensitive: bool = False) -> Optional[str]:
    """Search for a value inside a .docx file."""
    if Document is None:
        return None
    
    try:
        doc = Document(str(file_path))
        matches = []
        search_term = search_value if case_sensitive else search_value.lower()
        
        # Search in paragraphs
        for i, para in enumerate(doc.paragraphs, 1):
            text = para.text
            compare_text = text if case_sensitive else text.lower()
            if search_term in compare_text:
                preview = text[:50] + "..." if len(text) > 50 else text
                matches.append(f"Paragraph {i}: '{preview}'")
        
        # Search in tables
        for table_idx, table in enumerate(doc.tables, 1):
            for row_idx, row in enumerate(table.rows):
                for cell_idx, cell in enumerate(row.cells):
                    text = cell.text
                    compare_text = text if case_sensitive else text.lower()
                    if search_term in compare_text:
                        matches.append(f"Table {table_idx}, Row {row_idx + 1}, Cell {cell_idx + 1}")
        
        if matches:
            return "; ".join(matches[:5]) + (f" (+{len(matches)-5} more)" if len(matches) > 5 else "")
        return None
        
    except Exception:
        return None


def search_in_file(file_path: Path, search_value: str, case_sensitive: bool = False) -> Optional[str]:
    """Search for a value in a file based on its extension."""
    extension = file_path.suffix.lower()
    
    if extension == '.xlsx':
        return search_in_xlsx(file_path, search_value, case_sensitive)
    elif extension == '.xls':
        return search_in_xls(file_path, search_value, case_sensitive)
    elif extension == '.docx':
        return search_in_docx(file_path, search_value, case_sensitive)
    
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
        
        print(f"Scanning folder: {root_path}")
        
        # Walk through all directories
        for folder in root.rglob('*'):
            if folder.is_dir():
                # Skip 'Old' folders and their contents
                if is_old_folder(folder):
                    continue
                # Check if any parent is 'Old'
                skip = False
                for parent in folder.parents:
                    if is_old_folder(parent):
                        skip = True
                        break
                if skip:
                    continue
                
                # Check if this is a final folder
                if is_final_folder(folder):
                    final_folders.add(folder)
        
        # Also check if root itself is a final folder
        if is_final_folder(root):
            final_folders.add(root)
    
    return sorted(final_folders, key=lambda f: str(f).lower())


def search_in_final_folders(
    search_value: str,
    folder_paths: List[str],
    case_sensitive: bool = False
) -> List[FolderResult]:
    """
    Search for a value in final folders.
    
    For each final folder:
    - Lists all supported files
    - Searches only in the most recently modified file
    
    Args:
        search_value: The value to search for inside files.
        folder_paths: List of root folder paths to search in.
        case_sensitive: Whether the search should be case-sensitive.
    
    Returns:
        List of FolderResult objects.
    """
    
    results: List[FolderResult] = []
    
    # Find all final folders
    final_folders = find_all_final_folders(folder_paths)
    
    if not final_folders:
        print("\nNo final folders found.")
        return results
    
    print(f"\nFound {len(final_folders)} final folder(s) to process...")
    print()
    
    for idx, folder in enumerate(final_folders, 1):
        # Progress indicator
        print(f"Processing folder {idx}/{len(final_folders)}: {folder.name}", end='\r')
        
        # Get all supported files in this folder
        files = get_supported_files(folder)
        all_file_names = sorted([f.name for f in files])
        
        if not files:
            # No supported files in this folder
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
        
        # Find the most recently modified file
        most_recent = get_most_recent_file(files)
        
        if most_recent:
            # Get modification time
            mod_time = datetime.fromtimestamp(most_recent.stat().st_mtime)
            mod_time_str = mod_time.strftime("%Y-%m-%d %H:%M:%S")
            
            # Search in the most recent file
            search_result = search_in_file(most_recent, search_value, case_sensitive)
            
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
    
    print()  # New line after progress
    return results


def print_results(results: List[FolderResult], search_value: str) -> None:
    """Print search results in a formatted way."""
    if not results:
        print("\nNo folders processed.")
        return
    
    # Get only folders with matches
    folders_with_matches = [r for r in results if r.search_found]
    
    print(f"\n{'='*80}")
    print(f"SEARCH RESULTS FOR: '{search_value}'")
    print(f"{'='*80}")
    print(f"\nProcessed {len(results)} final folder(s)")
    print(f"Found matches in {len(folders_with_matches)} folder(s)")
    
    if not folders_with_matches:
        print("\nNo matches found in any folder.")
        return
    
    print(f"\n{'='*80}")
    print("FOLDERS WITH MATCHES:")
    print(f"{'='*80}")
    
    for i, result in enumerate(folders_with_matches, 1):
        print(f"\n{i}. FOLDER: {result.folder_name}")
        print(f"   Path: {result.folder_path}")
        print(f"   All files in folder ({len(result.all_files)}):")
        for fname in result.all_files:
            marker = " --> SEARCHED (most recent)" if fname == result.searched_file else ""
            print(f"      - {fname}{marker}")
        print(f"   Searched file: {result.searched_file} (modified: {result.searched_file_modified})")
        print(f"   MATCH FOUND: {result.search_details}")


# Main function
if __name__ == "__main__":
    print("=" * 70)
    print("File Content Search Utility")
    print("Searches in Excel (.xlsx, .xls) and Word (.docx) files")
    print("For each final folder, shows all files and searches the most recent one")
    print("Skips subfolders named 'Old'")
    print("=" * 70)
    print()
    
    if BASE_PATH:
        print(f"Base path: {BASE_PATH}")
        print("(You only need to enter folder names relative to this path)")
        print()
    
    # Check dependencies first
    if not check_dependencies():
        exit(1)
    
    # Get search value (keep asking until provided)
    search_value = ""
    while not search_value:
        search_value = input("Enter the value to search for: ").strip()
        if not search_value:
            print("Search value is required. Please try again.\n")
    
    # Get folder paths (one at a time, keep asking until at least one is provided)
    folder_paths = []
    while not folder_paths:
        if BASE_PATH:
            print(f"\nEnter folder names relative to {BASE_PATH}")
            print("(press Enter on empty line when done):")
        else:
            print("\nEnter full folder paths to search (press Enter on empty line when done):")
        path_num = 1
        
        while True:
            folder = input(f"  Folder {path_num}: ").strip()
            if not folder:
                break
            # Combine with BASE_PATH if set
            if BASE_PATH and not folder.startswith("/") and not folder.startswith("C:") and not folder.startswith("D:"):
                full_path = f"{BASE_PATH}/{folder}"
            else:
                full_path = folder
            folder_paths.append(full_path)
            path_num += 1
        
        if not folder_paths:
            print("At least one folder path is required. Please try again.")
    
    print(f"\nSearching in {len(folder_paths)} folder(s):")
    for fp in folder_paths:
        print(f"  - {fp}")
    
    try:
        # Perform search
        results = search_in_final_folders(
            search_value=search_value,
            folder_paths=folder_paths,
            case_sensitive=False
        )
        
        # Print results
        print_results(results, search_value)
        
    except Exception as e:
        print(f"An error occurred: {e}")
