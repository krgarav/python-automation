import os
import shutil

def cleanup_downloaded_spaces(base_path="downloaded_spaces", exclude_folder=None):
    """
    Deletes all files/folders inside 'downloaded_spaces' except one folder.

    Args:
        base_path (str): The parent folder containing downloaded space folders.
        exclude_folder (str): A folder name to keep (do NOT delete).
    """

    if not os.path.exists(base_path):
        print(f"⚠️ Folder '{base_path}' does not exist.")
        return

    for item in os.listdir(base_path):
        item_path = os.path.join(base_path, item)

        # Skip the folder that must remain
        if exclude_folder and item == exclude_folder:
            print(f"⏭️ Skipping (keeping): {item}")
            continue

        # Delete files and folders
        try:
            if os.path.isfile(item_path):
                os.remove(item_path)
                print(f"🗑️ Deleted file: {item}")
            else:
                shutil.rmtree(item_path)
                print(f"🗑️ Deleted folder: {item}")
        except Exception as e:
            print(f"❌ Error deleting {item}: {e}")

    print("✅ Cleanup completed!")
