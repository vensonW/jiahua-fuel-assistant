import os
from pathlib import Path
import sys

bundle = Path(sys._MEIPASS)
os.environ["TCL_LIBRARY"] = str(bundle / "_tcl_data")
os.environ["TK_LIBRARY"] = str(bundle / "_tk_data")
