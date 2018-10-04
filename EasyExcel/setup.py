import setuptools

with open("README.md", "r") as fh:
    long_description = fh.read()

setuptools.setup(
    name="EasyExcel",
    version="0.0.1",
    author="Josh Humphrey",
    author_email="jhumphrey321@gmail.com",
    description="A small library to interact with python and excel",
    long_description=long_description,
    long_description_content_type="text/markdown",
    url="https://github.com/JoshHumpherey/EasyExcel",
    packages=setuptools.find_packages(),
    classifiers=[
        "Programming Language :: Python :: 3",
        "License :: OSI Approved :: MIT License",
        "Operating System :: OS Independent",
    ],
)
