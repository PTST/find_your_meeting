from distutils.core import setup
from Cython.Build import cythonize

setup(
  name = 'find_room',
  ext_modules = cythonize(["*.pyx"]),
)