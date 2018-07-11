from distutils.core import setup
import py2exe

setup(console=[{'script':'coal_extraction.py'}],
      options={"py2exe":{"includes":["xlrd", "xlsxwriter"]}})
