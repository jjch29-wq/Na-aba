with open('home/src/Material-Master-Manager-V13.py', 'r', encoding='utf-8') as f:
    text = f.read()

text = text.replace("'layout_locked': self.layout_locked,", "'layout_locked': getattr(self, 'layout_locked', False),")
text = text.replace("'resolution_locked': self.resolution_locked,", "'resolution_locked': getattr(self, 'resolution_locked', False),")

with open('home/src/Material-Master-Manager-V13.py', 'w', encoding='utf-8') as f:
    f.write(text)
