from distutils.core import setup
import py2exe
import os
import sys

# Find GTK+ installation path
__import__('gtk')
m = sys.modules['gtk']
gtk_base_path = m.__path__[0]

setup(
    name = 'dipper-std',
    description = 'dipper-std',
    version = '1.0',

    windows = [
                  {
                      'script': 'dipper-std.py',
                  }
              ],

    options = {
                  'py2exe': {
                      'packages':'encodings',
                      # Optionally omit gio, gtk.keysyms, and/or rsvg if you're not using them
                      'includes': 'cairo, pango, pangocairo, atk, gobject, gio, gtk.keysyms, rsvg',
                  }
              },

    data_files=[
        'gui.glade',
        'select_dialog.glade',
        'select_sheet_dialog.glade',
        'miniscale.png',
    ]
)
