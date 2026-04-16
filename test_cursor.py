import tkinter as tk
from tkinter import ttk

root = tk.Tk()

cb = ttk.Combobox(root, values=['Apple', 'Avocado', 'Banana', 'Blueberry'])
cb.pack()

def on_key(event):
    if event.keysym in ('Up', 'Down', 'Left', 'Right', 'Return'): return
    
    cb['values'] = ['Apple', 'Avocado']
    cb.tk.call('ttk::combobox::Post', cb._w)
    
    # Try restoring focus
    cb.focus_set()
    cb.icursor(tk.END)
    print("Post called and focus_set called.")

cb.bind('<KeyRelease>', on_key)

root.after(500, lambda: [cb.focus_set(), cb.insert(0, 'a'), cb.event_generate('<KeyRelease>', keysym='a'), root.after(500, root.destroy)])
root.mainloop()
