import os
import argparse
import csv
from datetime import datetime
import win32security
import concurrent.futures
from queue import Queue

EXCLUDED_FOLDERS = {"$RECYCLE.BIN", "System Volume Information"}

# Define permission constants that map to the raw numbers
PERMISSION_MAP = {
    1: "Read",
    2: "Write",
    4: "Execute",
    16: "ReadAndExecute",
    32: "WriteAndExecute",
    256: "ReadAttributes",
    512: "WriteAttributes",
    8192: "Delete",
    32768: "ReadPermissions",
    131072: "ChangePermissions",
    262144: "TakeOwnership",
    16777216: "FullControl",
}

def format_size(size_in_bytes):
    """Format size in bytes into a human-readable format (KB, MB, GB)."""
    if size_in_bytes is None:
        return "0 B"
    elif size_in_bytes < 1024:
        return f"{size_in_bytes} B"
    elif size_in_bytes < 1024 ** 2:
        return f"{size_in_bytes / 1024:.2f} KB"
    elif size_in_bytes < 1024 ** 3:
        return f"{size_in_bytes / (1024 ** 2):.2f} MB"
    else:
        return f"{size_in_bytes / (1024 ** 3):.2f} GB"

def get_permissions(path):
    """Get human-readable permissions of a file or directory."""
    try:
        security_info = win32security.GetFileSecurity(path, win32security.DACL_SECURITY_INFORMATION)
        dacl = security_info.GetSecurityDescriptorDacl()
        permissions = set()  # Use a set to avoid duplicates
        for i in range(dacl.GetAceCount()):
            ace = dacl.GetAce(i)
            access_mask = ace[1]
            for perm, perm_value in PERMISSION_MAP.items():
                if access_mask & perm:
                    permissions.add(perm_value)  # Add to set, duplicates will be ignored
        return ", ".join(sorted(permissions)) if permissions else "No Permissions"
    except Exception as e:
        print(f"Error retrieving permissions for {path}: {e}")
        return "Error retrieving permissions"


def get_owner(path):
    """Get the owner of the file or directory."""
    try:
        security_info = win32security.GetFileSecurity(path, win32security.OWNER_SECURITY_INFORMATION)
        owner_sid = security_info.GetSecurityDescriptorOwner()
        owner_name, domain, _ = win32security.LookupAccountSid(None, owner_sid)
        return f"{domain}\\{owner_name}"
    except Exception as e:
        print(f"Error retrieving owner for {path}: {e}")
        return "Error retrieving owner"



def get_file_info(path):
    """Retrieve file or directory information, excluding certain system directories."""
    try:
        stats = os.stat(path)
        is_directory = os.path.isdir(path)
        size = stats.st_size if not is_directory else None
        extension = None if is_directory else os.path.splitext(path)[1]
        file_created = datetime.fromtimestamp(stats.st_birthtime)
        permissions = get_permissions(path)
        owner = get_owner(path)

        # Initialize the item counts to None for files
        number_of_items = None
        file_count = None
        folder_count = None

        if is_directory and os.path.basename(path) not in EXCLUDED_FOLDERS:
            # Count files and folders in the directory
            file_count = 0
            folder_count = 0
            for entry in os.scandir(path):
                if entry.is_file():
                    file_count += 1
                elif entry.is_dir() and entry.name not in EXCLUDED_FOLDERS:
                    folder_count += 1
            number_of_items = file_count + folder_count  # Sum of files and folders

            # Calculate total size for directories
            size = get_directory_size(path)

        return {
            'Name': os.path.basename(path),
            'Path': os.path.abspath(path),
            'Type': 'Directory' if is_directory else 'File',
            'Extension': extension,
            'Size': format_size(size),
            'Permissions': permissions,
            'Owner': owner,
            'CreatedDate': file_created,
            'ModifiedDate': datetime.fromtimestamp(stats.st_mtime),
            'NumberOfItems': number_of_items,  # Will be None for files
            'FolderCount': folder_count,  # Will be None for files
            'FileCount': file_count  # Will be None for files
        }
    except Exception as e:
        print(f"Error retrieving information for {path}: {e}")
        return None

def get_directory_size(path):
    """Calculate the total size of a directory (excluding certain system directories)."""
    total_size = 0
    try:
        for entry in os.scandir(path):
            if entry.name in EXCLUDED_FOLDERS:
                continue  # Skip excluded folders
            if entry.is_file():
                total_size += entry.stat().st_size
            elif entry.is_dir():
                total_size += get_directory_size(entry.path)  # Recursively add size of subdirectories
    except PermissionError as e:
        print(f"Permission error accessing {path}: {e}")
    except Exception as e:
        print(f"Error calculating size for {path}: {e}")
    return total_size

def scan_directory(path, queue):
    """Recursively scan the directory and gather file information in parallel."""
    if path.endswith(":"):
        path = path + "\\"  # Append backslash for full drive scan

    root_info = get_file_info(path)
    if root_info:
        root_size = get_directory_size(path)
        root_info['Size'] = format_size(root_size)
        queue.put(root_info)

    with concurrent.futures.ThreadPoolExecutor(max_workers=100) as executor:
        futures = []
        for root, dirs, files in os.walk(path, topdown=True):
            dirs[:] = [d for d in dirs if d not in EXCLUDED_FOLDERS]
            for file in files:
                file_path = os.path.join(root, file)
                futures.append(executor.submit(process_file, file_path, queue))
            for directory in dirs:
                dir_path = os.path.join(root, directory)
                futures.append(executor.submit(process_directory, dir_path, queue))

        for future in concurrent.futures.as_completed(futures):
            pass


def process_file(file_path, queue):
    """Process a file and return its info."""
    file_info = get_file_info(file_path)
    if file_info:
        queue.put(file_info)

def process_directory(dir_path, queue):
    """Process a directory and return its info."""
    dir_info = get_file_info(dir_path)
    if dir_info:
        dir_size = get_directory_size(dir_path)
        dir_info['Size'] = format_size(dir_size)
        queue.put(dir_info)

def save_to_csv(file_info_list, output_csv):
    """Save the collected file information to a CSV file."""
    try:
        # Sort file_info_list by the 'Path' field in ascending order
        file_info_list = sorted(file_info_list, key=lambda x: x['Path'].lower())

        with open(output_csv, mode='w', newline='', encoding='utf-8') as csvfile:
            fieldnames = ['Name', 'Path', 'Type', 'Extension',   'CreatedDate', 'ModifiedDate','Permissions','Owner', 'Size', 'NumberOfItems',  'FolderCount','FileCount']
            writer = csv.DictWriter(csvfile, fieldnames=fieldnames)

            writer.writeheader()
            for file_info in file_info_list:
                # Convert FileCreated and FileModified to date and time format
                file_info['CreatedDate'] = file_info['CreatedDate'].strftime('%Y-%m-%d %H:%M:%S')
                file_info['ModifiedDate'] = file_info['ModifiedDate'].strftime('%Y-%m-%d %H:%M:%S')
                writer.writerow(file_info)
        print(f"File discovery completed. Output saved to: {output_csv}")
    except Exception as e:
        print(f"Error: Unable to save the output CSV file. {e}")

if __name__ == "__main__":
    start_time = datetime.now()  # Capture the start time
    parser = argparse.ArgumentParser(description="File Discovery Script")
    parser.add_argument("searchPath", help="Path to search (e.g., D: or C:\\path\\to\\dir)")
    parser.add_argument("outputDir", help="Directory to save the output CSV file")

    args = parser.parse_args()

    if not args.outputDir.endswith("\\"):
        args.outputDir += "\\"

    current_date = datetime.now().strftime("%Y%m%d_%H%M%S")
    output_csv = f"{args.outputDir}FileDiscovery_{current_date}.csv"

    # Using Queue for thread-safe collection of results
    queue = Queue()

    scan_directory(args.searchPath, queue)

    # Collect all results from the queue
    file_info_list = []
    while not queue.empty():
        file_info_list.append(queue.get())

    save_to_csv(file_info_list, output_csv)

    end_time = datetime.now()  # Capture the end time
    elapsed_time = end_time - start_time  # Subtract the start time from the end time
    # Print the elapsed time
    print(f"Time taken for the scan: {elapsed_time}")
