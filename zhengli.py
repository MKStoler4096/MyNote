import os
import shutil

def process_folder(folder_path):
    for subfolder_name in os.listdir(folder_path):
        subfolder_path = os.path.join(folder_path, subfolder_name)
        if os.path.isdir(subfolder_path):
            folder_names = subfolder_name.split("-")
            if len(folder_names) == 1:
                print(f"Skipping folder: {subfolder_path}")
                continue
            new_path = folder_path
            for folder_name in folder_names:
                new_path = os.path.join(new_path, folder_name)
                if not os.path.exists(new_path):
                    os.makedirs(new_path)
                    print(f"Created folder: {new_path}")
            if os.path.exists(new_path):
                shutil.move(subfolder_path, new_path)
                print(f"Moved files from {subfolder_path} to {new_path}")
            move_sub_folders_and_files(subfolder_path, new_path)
            os.rmdir(subfolder_path)
            print(f"Deleted folder: {subfolder_path}")

def move_sub_folders_and_files(source_folder, destination_folder):
    for item_name in os.listdir(source_folder):
        item_path = os.path.join(source_folder, item_name)
        if os.path.isdir(item_path):
            if not os.path.exists(os.path.join(destination_folder, item_name)):
                os.makedirs(os.path.join(destination_folder, item_name))
            move_sub_folders_and_files(item_path, os.path.join(destination_folder, item_name))
        elif os.path.isfile(item_path):
            if not os.path.exists(os.path.join(destination_folder, item_name)):
                shutil.move(item_path, destination_folder)
                print(f"Moved file: {item_path} to {destination_folder}")

def show_folder_structure(folder_path, indent=""):
    print(indent + os.path.basename(folder_path))
    for item_name in os.listdir(folder_path):
        item_path = os.path.join(folder_path, item_name)
        if os.path.isdir(item_path):
            show_folder_structure(item_path, indent + "    ")

source_folder = input("Select the folder to be organized: ")
show_folder_structure(source_folder)
process_folder(source_folder)
show_folder_structure(source_folder)
