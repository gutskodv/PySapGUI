from setuptools import setup
import pysapgui

with open('requirements.txt') as f:
    required = f.read().splitlines()

with open("README.md", "r") as fh:
    long_description = fh.read()

setup(
    name='PySapGUI',
    version=pysapgui.__version__,
    packages=['pysapgui'],
    url='https://github.com/gutskodv/PySapGUI',
    license='GPL v2',
    author='Dmitry Gutsko',
    author_email='gutskodv@gmail.com',
    description='SAP GUI Scripting Library in Python',
    long_description_content_type="text/markdown",
    long_description=long_description,
    install_requires=required,
    include_package_data=True
)
