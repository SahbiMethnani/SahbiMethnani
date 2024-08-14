# setup.py

from setuptools import setup, find_packages

setup(
    name='SahbiMethnani',
    version='0.1',
    packages=find_packages(),
    install_requires=[
        'PyQt5',
        'openpyxl',
        'sqlite3',
    ],
    author='sahbimethnani',
    author_email='sahbimethnani3@gmail.com',
    description='Une bibliothèque pour la gestion des employés avec une interface graphique utilisant PyQt5 et SQLite.',
    url='https://github.com/sahbimethnani/SahbiMethnani',
    classifiers=[
        'Programming Language :: Python :: 3',
        'License :: OSI Approved :: MIT License',
        'Operating System :: OS Independent',
    ],
    python_requires='>=3.6',
)
