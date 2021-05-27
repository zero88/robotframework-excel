import sys
import re
import pip
from setuptools import setup, find_packages
from os.path import abspath, join, dirname


_PACKAGE = 'ExcelRobot'

sys.path.append(join(dirname(__file__), _PACKAGE))

_version_path = join(dirname(__file__), _PACKAGE, 'version.py')
with open(_version_path) as f:
    code = compile(f.read(), _version_path, 'exec')
    exec(code)

_DESCRIPTION = """
This test library provides some keywords to allow
opening, reading, writing, and saving Excel files
from Robot Framework.
"""[1:-1]

_URL = 'https://github.com/zero-88/robotframework-excel'
_DOWNLOAD_URL = _URL + '/tarball/' + VERSION

def __gather_dependencies(require_file):
    with open(join(dirname(abspath(__file__)), 'requirements.txt')) as f:
        _reqs = f.read().splitlines()
    _links = []
    return _reqs, _links


_REQUIRES, _LINKS, = __gather_dependencies('requirements.txt')

setup(
    name='robotframework-excel',
    version=VERSION,
    description='Robot Framework',
    long_description=_DESCRIPTION,
    author='zero',
    author_email='<sontt246@gmail.com>',
    url=_URL,
    license='Apache License 2.0',
    keywords='robotframework testing testautomation excel',
    platforms='any',
    python_requires='>=2.7, !=3.0.*, !=3.1.*, !=3.2.*, <4',
    classifiers=[
        "License :: OSI Approved :: Apache Software License",
        "Development Status :: 5 - Production/Stable",
        "Programming Language :: Python",
        "Intended Audience :: Developers",
        "Programming Language :: Python :: 2.7",
        "Programming Language :: Python :: 3",
        "Programming Language :: Python :: 3.3",
        "Programming Language :: Python :: 3.4",
        "Programming Language :: Python :: 3.5",
        "Programming Language :: Python :: 3.6",
        "Topic :: Software Development :: Testing",
        "Topic :: Software Development :: Quality Assurance"
    ],
    install_requires=_REQUIRES,
    dependency_links=_LINKS,
    packages=find_packages(exclude=['tests']),
    data_files=[('ExcelRobotTest', ['docs/ExcelRobot.html', 'docs/release-notes.md'])],
    download_url=_DOWNLOAD_URL,
)
