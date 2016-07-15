
try:
    from setuptools import setup
except ImportError:
    from distutils.core import setup

config = {
    'description': 'Make test with questions about documents',
    'author': 'Jorge Soto Garcia',
    'url': 'http://www.github.com/sotogarcia',
    'download_url': 'http://www.github.com/sotogarcia',
    'author_email': 'sotogarcia@gmail.com',
    'version': '0.1',
    'install_requires': [],
    'packages': ['mktest'],
    'scripts': [],
    'name': 'mktest'
}

setup(**config)
