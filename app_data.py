import json
import os
from pathlib import Path

class AppData:
    def __init__(self):
        self.data_file = 'app_settings.json'
        data = self.load_data()
        self.folders = data.get('folders', [])
        self.favorites = data.get('favorites', [])
        self.last_script_count = data.get('last_script_count', 0)
        
    def load_data(self):
        try:
            if os.path.exists(self.data_file):
                with open(self.data_file, 'r') as f:
                    return json.load(f)
        except Exception as e:
            print(f"Error loading data: {e}")
        return {'folders': [], 'favorites': []}
        
    def save_data(self):
        try:
            with open(self.data_file, 'w') as f:
                json.dump({
                    'folders': self.folders,
                    'favorites': self.favorites,
                    'last_script_count': self.last_script_count
                }, f)
        except Exception as e:
            print(f"Error saving data: {e}")
            
    def update_script_count(self, count):
        self.last_script_count = count
        self.save_data()
            
    def add_folder(self, folder_path):
        if folder_path not in self.folders:
            self.folders.append(folder_path)
            self.save_data()
            
    def remove_folder(self, folder_path):
        if folder_path in self.folders:
            self.folders.remove(folder_path)
            self.save_data()
            
    def toggle_favorite(self, script_path):
        if script_path in self.favorites:
            self.favorites.remove(script_path)
        else:
            self.favorites.append(script_path)
        self.save_data()
        
    def is_favorite(self, script_path):
        return script_path in self.favorites
            
    def get_all_powershell_scripts(self):
        scripts = []
        for folder in self.folders:
            if os.path.exists(folder):
                for root, _, files in os.walk(folder):
                    for file in files:
                        if file.endswith('.ps1'):
                            full_path = os.path.join(root, file)
                            relative_path = os.path.relpath(full_path, folder)
                            scripts.append({
                                'name': file,
                                'full_path': full_path,
                                'relative_path': relative_path,
                                'folder': folder,
                                'is_favorite': self.is_favorite(full_path)
                            })
        return scripts
