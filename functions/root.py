import os
import sys

# Si el script está corriendo como un ejecutable, usar la ruta temporal _MEIPASS
if hasattr(sys, '_MEIPASS'):
    ROOT_BASE = os.path.join(sys._MEIPASS, "fixed")
else:
    # Si no está empaquetado, usar la ruta relativa habitual
    ROOT_BASE = os.path.join(os.getcwd(), "fixed")

ROOT_GDB = ROOT_BASE +  "/gdb/"
ROOT_RESULT = ROOT_BASE +  "/result/"
ROOT_IMG = ROOT_BASE +  "/images/"
