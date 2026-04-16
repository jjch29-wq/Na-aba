import os, sys, json

# Emulate the class paths
app_dir = os.path.dirname(os.path.abspath('Material-Master-Manager-V13.py'))
documents_dir = os.path.join(os.path.expanduser('~'), 'Documents')
config_path_default = os.path.join(documents_dir, 'Material_Manager_Config.json')

print(f"App Dir: {app_dir}")
print(f"Documents Dir: {documents_dir}")
print(f"Default Config Path: {config_path_default}")
print(f"Exists? {os.path.exists(config_path_default)}")

# Let's search for any JSON in Documents
if os.path.exists(documents_dir):
    print("\nFiles in Documents:")
    for f in os.listdir(documents_dir):
        if f.endswith('.json'):
            print(f" - {f}")
            if 'Config' in f:
                with open(os.path.join(documents_dir, f), 'r', encoding='utf-8') as jf:
                    data = json.load(jf)
                    print(f"   budget_view_custom_columns: {data.get('budget_view_custom_columns', [])}")

# Check the app directory too
print("\nFiles in App Dir:")
for f in os.listdir(app_dir):
    if f.endswith('.json'):
        print(f" - {f}")
