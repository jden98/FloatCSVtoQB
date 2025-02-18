from py2exe import freeze
py2exe_options = dict(
                ascii=True,  # Exclude encodings
                excludes=['_ssl',  # Exclude _ssl
                        'pyreadline', 'difflib', 'doctest', 'locale', 
                        'optparse', 'calendar',
                        "Tkconstants","Tkinter","tcl"],  # Exclude standard library
                compressed=True,  # Compress library.zip
                )

freeze(console=['Float2QB.py'],
       options={'py2exe': py2exe_options}) 
