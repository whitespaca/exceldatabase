from setuptools import setup, find_packages

setup(
    name="exceldatabase",
    version="0.1.0",
    packages=find_packages(),
    install_requires=[
        "openpyxl"
    ],
    author="whitespaca",
    author_email="whitespaca@outlook.com",
    description="A simple Excel-based database system using openpyxl.",
    long_description=open("README.md").read(),
    long_description_content_type="text/markdown",
    url="https://github.com/whitespaca/exceldatabase",
    classifiers=[
        "Programming Language :: Python :: 3",
        "License :: OSI Approved :: MIT License",
        "Operating System :: OS Independent",
    ],
    python_requires=">=3.7",
)