import traceback
import tkinter as tk
import sys
import os

sys.path.append(os.path.join(os.getcwd(), 'home', 'src'))
from importlib.util import spec_from_file_location, module_from_spec

spec = spec_from_file_location("main_app", "home/src/Material-Master-Manager-V13.py")
mod = module_from_spec(spec)
spec.loader.exec_module(mod)

root = tk.Tk()
try:
    app = mod.MaterialManager(root)
    # Simulate clicking In/Out tab
    app.notebook.select(app.tab_inout)
    app.on_tab_changed(None)
    print("SUCCESS")
except Exception as e:
    print("ERROR:")
    traceback.print_exc()
finally:
    root.destroy()
