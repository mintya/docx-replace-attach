import setuptools
from setuptools import setup

with open("README.md", "r") as fh:
    long_description = fh.read()

setup(
    name="docx-replace-attach",
    version="0.0.1",
    author="mintya",
    author_email="931108724@qq.com",
    description="Replace key words to attachments inside a document of MS Word",
    long_description=long_description,
    long_description_content_type="text/markdown",
    url="https://github.com/mintya/docx-replace-attach",
    packages=setuptools.find_packages(),
    install_requires=['python-docx>=1.1.0'],
    entry_points={},
    classifiers=[
        "Programming Language :: Python :: 3",
        "License :: OSI Approved :: MIT License",
        "Operating System :: OS Independent",
    ],
)