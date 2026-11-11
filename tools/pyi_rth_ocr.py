"""Initialize ONNX before PyInstaller's Qt hook in frozen OCR child processes."""

import sys

if '--multiprocessing-fork' in sys.argv:
    # Qt5 ships an older C++ runtime; ONNX must initialize before QtCore loads.
    import onnxruntime
