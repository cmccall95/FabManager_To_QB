import os

def get_folder_size(folder_path):
    print("Running get_folder_size")
    total_size = 0
    for dirpath, dirnames, filenames in os.walk(folder_path):
        for filename in filenames:
            file_path = os.path.join(dirpath, filename)
            try:
                total_size += os.path.getsize(file_path)
            except (OSError, FileNotFoundError):
                continue
    
    return total_size / (1024 * 1024 * 1024)  # Convert bytes to GB

if __name__ == '__main__':
    path1 = r"S:\Shared With Me\23 Fabrication\Fab QC\30489 OCI\6. Document Package" # 25.41
    path2 = r"S:\Shared With Me\23 Fabrication\Fab QC\30501 AQUATECH (OCI)"

    size_gb_1 = get_folder_size(path1)
    size_gb_2 = get_folder_size(path2)
    total_size = size_gb_1 + size_gb_2
    print(f"Folder size: {total_size:.2f} GB")
