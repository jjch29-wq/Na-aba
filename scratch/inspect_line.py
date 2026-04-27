file_path = r"c:\Users\-\OneDrive\바탕 화면\PMI Report\home\src\Material-Master-Manager-V13.py"
with open(file_path, 'rb') as f:
    lines = f.readlines()

start_line = 11889
end_line = 11892
for target_line in range(start_line, end_line + 1):
    if len(lines) >= target_line:
        line = lines[target_line - 1]
        print(f"Line {target_line} hex: {line.hex()}")
        try:
            print(f"Line {target_line} utf8: {line.decode('utf-8')}")
        except:
            print(f"Line {target_line} utf8: failed")
    else:
        print(f"File only has {len(lines)} lines")
