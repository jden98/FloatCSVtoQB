from py2exe import freeze
import click

py2exe_options = {
                "ascii": True,  # Exclude encodings
                "packages": ["click"],
                "includes": ["click", "click.core", "click.decorators", "click.parser", "click.termui", "click.types"],
                "excludes": ["_ssl",  # Exclude _ssl
                        'pyreadline', 'difflib', 'doctest', 'optparse', 
                        "Tkconstants","Tkinter","tcl"],  # Exclude some standard libraries
                "compressed": True,  # Compress library.zip
                }

freeze(console=['Float2QB.py'],
       options={'py2exe': py2exe_options}) 