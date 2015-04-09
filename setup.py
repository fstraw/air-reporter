try:
    from setuptools import setup
except ImportError:
    from distutils.core import setup

config = {
    'description': 'air-reporting',
    'author': 'Brandon Batt',
    'url': 'https://github.com/fstraw/air-reporter',
    'download_url': 'https://github.com/fstraw/air-reporter',
    'author_email': 'fstraw@lowestfrequency.com',
    'version': '0.2',
    'install_requires': ['python-docx>=0.8.5'],
    'packages': [''],
    'scripts': [],
    'name': 'Air Reporting'
}

setup(**config)