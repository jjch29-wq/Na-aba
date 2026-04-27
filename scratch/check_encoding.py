import sys

file_path = r"c:\Users\-\OneDrive\바탕 화면\PMI Report\home\src\Material-Master-Manager-V13.py"

try:
    with open(file_path, 'rb') as f:
        content = f.read()
    
    # Try decoding as CP949
    decoded = content.decode('cp949')
    lines = decoded.splitlines()
    
    start = 11880
    end = 11900
    for i in range(start, min(end, len(lines))):
        print(f"{i+1}: {lines[i]}")

except Exception as e:
    print(f"Error: {e}")
