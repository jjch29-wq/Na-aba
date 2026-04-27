import os

file_path = r"c:\Users\-\OneDrive\바탕 화면\PMI Report\home\src\Material-Master-Manager-V13.py"
backup_path = file_path + ".bak"

# 1. Read the file as bytes
with open(file_path, 'rb') as f:
    content = f.read()

# 2. Backup
with open(backup_path, 'wb') as f:
    f.write(content)

# 3. Identify the broken sequence. 
# Line 11889 hex starts with 80 ec 9e a5
# The full line from my check was:
# 80 ec 9e a5 22 2c 20 63 6f 6d 6d 61 6e 64 3d 5f 73 61 76 65 2c 20 77 69 64 74 68 3d 31 30 29 2e 70 61 63 6b 28 73 69 64 65 3d 27 6c 65 66 74 27 2c 20 70 61 64 78 3d 38 29 0d 0a
broken_bytes = bytes.fromhex("80ec9ea5222c20636f6d6d616e643d5f736176652c2077696474683d3130292e7061636b28736964653d276c656674272c20706164783d38290d0a")

# The correct sequence for "저장" is eca080ec9ea5
# So the line should be:
fixed_bytes = bytes.fromhex("eca080ec9ea5222c20636f6d6d616e643d5f736176652c2077696474683d3130292e7061636b28736964653d276c656674272c20706164783d38290d0a")

# Also, let's add the encoding header if it's missing.
header = b"# -*- coding: utf-8 -*-\n"
if not content.startswith(b"# -*- coding: utf-8 -*-") and not content.startswith(b"### VERSION"):
    # If it starts with VERSION, maybe put it after that
    parts = content.split(b"\n", 1)
    if b"VERSION" in parts[0]:
        new_content = parts[0] + b"\n" + header + parts[1]
    else:
        new_content = header + content
else:
    new_content = content

# Replace the broken line
# We look for the exact broken bytes to be sure.
if broken_bytes in new_content:
    new_content = new_content.replace(broken_bytes, fixed_bytes)
    print("Successfully replaced the broken line.")
else:
    print("Could not find the exact broken byte sequence. Trying a partial match.")
    # Try just the start
    partial_broken = bytes.fromhex("80ec9ea5222c20636f6d6d616e643d5f73617665")
    partial_fixed = bytes.fromhex("eca080ec9ea5222c20636f6d6d616e643d5f73617665")
    if partial_broken in new_content:
        new_content = new_content.replace(partial_broken, partial_fixed)
        print("Successfully replaced the partial broken match.")
    else:
        print("Partial match also failed.")

# Write back
with open(file_path, 'wb') as f:
    f.write(new_content)

print("Done.")
