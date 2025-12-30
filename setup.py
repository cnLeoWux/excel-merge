from setuptools import setup, find_packages

with open("README.md", "r", encoding="utf-8") as fh:
    long_description = fh.read()

with open("requirements.txt", "r", encoding="utf-8") as fh:
    requirements = fh.read().splitlines()

setup(
    name="excel-merge-tool",
    version="1.0.0",
    author="",
    author_email="",
    description="A Python tool to merge Excel files based on specific business logic",
    long_description=long_description,
    long_description_content_type="text/markdown",
    url="https://github.com/cnLeoWux/excel-merge",
    packages=find_packages(),
    include_package_data=True,
    install_requires=requirements,
    classifiers=[
        "Programming Language :: Python :: 3",
        "License :: OSI Approved :: MIT License",
        "Operating System :: OS Independent",
    ],
    python_requires='>=3.7',
    entry_points={
        'console_scripts': [
            'excel-merge=excel_merge:main',
            'excel-merge-cli=cli:main_cli',
        ],
    },
)