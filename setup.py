from distutils.core import setup

setup(name='SMS-Fisher',
      version='1.0',
      description='Converts Santini bills and catalogs to excel sheets for import into Fishbowl.',
      author='Amani Medcroft',
      author_email='amani@santini-us.com',
      url='https://github.com/slurpinpuffs/SMS-Fisher',
      install_requires=['openpyxl>=3.1.2', 'PyPDF2>=3.0.1'],
     )