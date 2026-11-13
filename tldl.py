"""Superseded: the shared engine (docx_translate.py) is fast and keeps formatting,
so this script now just runs the command-line translator in op2.py.
Usage is the same as op2.py:  python tldl.py [file.docx] [--lang fr]
"""
from op2 import main

if __name__ == "__main__":
    main()
