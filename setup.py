from setuptools import setup

kwargs = {'author': 'Metamorph',
 'author_email': '',
 'classifiers': ['Intended Audience :: Science/Research',
                 'Topic :: Scientific/Engineering'],
 'description': 'OpenMDAO Wrapper for Excel',
 'download_url': '',
 'entry_points': '[openmdao.component]\nexcel_wrapper.excel_wrapper.ExcelWrapper=excel_wrapper.excel_wrapper:ExcelWrapper\n\n[openmdao.container]\nexcel_wrapper.excel_wrapper.ExcelWrapper=excel_wrapper.excel_wrapper:ExcelWrapper',
 'include_package_data': True,
 'install_requires': ['openmdao'],
 'keywords': ['openmdao, excel'],
 'license': 'GNU General Public License, version 2',
 'maintainer': 'Young-Ki Lee',
 'maintainer_email': 'ylee@asdl.gatech.edu',
 'name': 'excel_wrapper',
 'package_data': {'excel_wrapper': ['sphinx_build/html/.buildinfo',
                                    'sphinx_build/html/genindex.html',
                                    'sphinx_build/html/index.html',
                                    'sphinx_build/html/objects.inv',
                                    'sphinx_build/html/pkgdocs.html',
                                    'sphinx_build/html/py-modindex.html',
                                    'sphinx_build/html/search.html',
                                    'sphinx_build/html/searchindex.js',
                                    'sphinx_build/html/srcdocs.html',
                                    'sphinx_build/html/usage.html',
                                    'sphinx_build/html/_modules/index.html',
                                    'sphinx_build/html/_modules/excel_wrapper/test/test_excel_wrapper.html',
                                    'sphinx_build/html/_sources/index.txt',
                                    'sphinx_build/html/_sources/pkgdocs.txt',
                                    'sphinx_build/html/_sources/srcdocs.txt',
                                    'sphinx_build/html/_sources/usage.txt',
                                    'sphinx_build/html/_static/ajax-loader.gif',
                                    'sphinx_build/html/_static/basic.css',
                                    'sphinx_build/html/_static/comment-bright.png',
                                    'sphinx_build/html/_static/comment-close.png',
                                    'sphinx_build/html/_static/comment.png',
                                    'sphinx_build/html/_static/default.css',
                                    'sphinx_build/html/_static/doctools.js',
                                    'sphinx_build/html/_static/down-pressed.png',
                                    'sphinx_build/html/_static/down.png',
                                    'sphinx_build/html/_static/file.png',
                                    'sphinx_build/html/_static/jquery.js',
                                    'sphinx_build/html/_static/minus.png',
                                    'sphinx_build/html/_static/plus.png',
                                    'sphinx_build/html/_static/pygments.css',
                                    'sphinx_build/html/_static/searchtools.js',
                                    'sphinx_build/html/_static/sidebar.js',
                                    'sphinx_build/html/_static/underscore.js',
                                    'sphinx_build/html/_static/up-pressed.png',
                                    'sphinx_build/html/_static/up.png',
                                    'sphinx_build/html/_static/websupport.js',
                                    'test/excel_wrapper_test.xlsx',
                                    'test/excel_wrapper_test.xml']},
 'package_dir': {'excel_wrapper': 'excel_wrapper'},
 'packages': ['excel_wrapper', 'excel_wrapper.test'],
 'url': '',
 'version': '0.5',
 'zip_safe': False}


setup(**kwargs)
