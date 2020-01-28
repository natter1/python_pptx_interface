import pathlib
from setuptools import setup

# The directory containing this file
root_path = pathlib.Path(__file__).parent
long_description = (root_path / "README.rst").read_text()

setup(
    name='python-pptx-interface',
    version='0.0.6.a02',
    packages=['pptx_tools'],
    url='https://github.com/natter1/python_pptx_interface.git',
    license='MIT',
    author='Nathanael Jöhrmann',
    author_email='',
    description='Easy interface to create pptx-files using python-pptx',
    long_description=long_description,
    long_description_content_type='text/x-rst',
    install_requires=[
        "python-pptx",
    ],
    package_data={
        'pptx_tools': ['resources/example-template.pptx']
    }
)