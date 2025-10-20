import os
import shutil

def file_extract_test(source_folder, destination_folder):

    # Check if the Destination folder exists
    os.makedirs(destination_folder, exist_ok=True)

    # Iterate over the contents of the source folder
    for folder2 in os.listdir(source_folder):                       
        folder2_path = os.path.join(source_folder, folder2)
        if not os.path.isdir(folder2_path):
            continue

        # Iterate over each subfolder in Source (Folder with Date Title)
        for folder3 in os.listdir(folder2_path):                    

            folder3_path = os.path.join(folder2_path, folder3)
            if not os.path.isdir(folder3_path):
                continue

            # Iterate over each subfolder in Date Folder (Folder with Computer Number)
            for folder4 in os.listdir(folder3_path):
                folder4_path = os.path.join(folder3_path, folder4)
                if not os.path.isdir(folder4_path):
                    continue

                #Create a Folder in the Destination Folder named with the Student ID
                dest_folder4 = os.path.join(destination_folder, folder4)

                #Copy Items in this subfolder (Folder Containg the Part files stored in each Student ID Folder)
                shutil.copytree(folder4_path, dest_folder4, dirs_exist_ok=True)
                print(f"Copied '{folder4_path}' to '{dest_folder4}'")


if __name__ == "__main__":
    # Ask the user for the source and destination folder paths.
    source = input("Enter the source folder: ").strip()
    destination = input("Enter the destination folder: ").strip()

    # Verify that the source folder exists.
    if not os.path.isdir(source):
        print("Source folder does not exist. Please check the path and try again.")
    else:
        file_extract_test(source, destination)
        print("Copy operation completed successfully!")