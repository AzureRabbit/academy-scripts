
try:
    from setuptools import setup
except ImportError:
    from distutils.core import setup

config = {
    'description': 'Sisgap scrapper',
    'author': 'Jorge Soto Garcia',
    'url': 'https://github.com/sotogarcia/academy-scripts',
    'download_url': 'https://github.com/sotogarcia/academy-scripts',
    'author_email': 'Jorge Soto Garcia.',
    'version': '0.1',
    'install_requires': ['nose'],
    'packages': ['sisgap'],
    'scripts': ['apiclient', 'oauth2client'],
    'name': 'Sisgap'
}

setup(**config)
