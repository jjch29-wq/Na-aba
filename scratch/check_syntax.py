import traceback
import sys

try:
    with open('Material-Master-Manager-V14.py', 'r', encoding='utf-8') as f:
        source = f.read()
    compile(source, 'Material-Master-Manager-V14.py', 'exec')
    print("Syntax OK")
except SyntaxError as e:
    print(f"SyntaxError: {e.msg}")
    print(f"File: {e.filename}")
    print(f"Line: {e.lineno}")
    print(f"Offset: {e.offset}")
    print(f"Text: {e.text}")
except Exception as e:
    traceback.print_exc()
