import os
import csv
from pathlib import Path
from datetime import datetime

def create_photo_catalog(directory):
    catalog = []
    for root, dirs, files in os.walk(directory):
        for file in files:
            if file.lower().endswith(('.jpg', '.jpeg', '.png', '.gif', '.bmp')):
                filepath = os.path.join(root, file)
                stat = os.stat(filepath)
                catalog.append({
                    'filename': file,
                    'path': filepath,
                    'size': stat.st_size,
                    'modified': datetime.fromtimestamp(stat.st_mtime).strftime('%Y-%m-%d %H:%M:%S')
                })
    return catalog

def save_to_csv(catalog, output_file):
    with open(output_file, 'w', newline='', encoding='utf-8') as csvfile:
        fieldnames = ['filename', 'path', 'size', 'modified']
        writer = csv.DictWriter(csvfile, fieldnames=fieldnames)
        writer.writeheader()
        for item in catalog:
            writer.writerow(item)

if __name__ == "__main__":
    directory = "."  # Current directory
    catalog = create_photo_catalog(directory)
    output_file = "photo_catalog.csv"
    save_to_csv(catalog, output_file)
    print(f"Photo catalog saved to {output_file}")