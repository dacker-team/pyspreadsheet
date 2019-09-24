from setuptools import setup, find_packages

with open('README.rst') as f:
    readme = f.read()

setup(
    name='pyspreadsheet',
    version='0.1.4',
    description='Easily send data to Google Sheets',
    long_description=readme,
    author='Dacker',
    author_email='hello@dacker.co',
    url='https://github.com/dacker-team/pyspreadsheet',
    keywords='send data google spreadsheet sheets easy',
    packages=find_packages(exclude=('tests', 'docs')),
    python_requires='>=3',
    install_requires=[
        "googleauthentication>=0.0.7",
        "dbstream>=0.0.6",
        "google-api-python-client>=1.6.6"
    ],
)
