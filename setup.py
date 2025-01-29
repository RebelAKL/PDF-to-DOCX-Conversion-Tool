from setuptools import setup, Extension
from Cython.Build import cythonize
import numpy  # Remove if not using numpy

extensions = [
    Extension("pdf_processor", 
              ["pdf_processor.pyx"],
              extra_compile_args=["-O3", "-march=native", "-ffast-math"],
              define_macros=[("NPY_NO_DEPRECATED_API", "NPY_1_7_API_VERSION")]
             )
]

setup(
    ext_modules=cythonize(extensions, compiler_directives={'binding': True}),
    include_dirs=[numpy.get_include()],  # Remove if not using numpy
    install_requires=[
        'PyPDF2>=3.0.0',
        'pdf2docx>=0.5.6',
        'python-docx>=0.8.11',
        'Cython>=0.29.32'
    ]
)