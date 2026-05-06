from setuptools import setup, find_packages

setup(
    name="cli-anything-onlyoffice",
    version="4.4.16",
    description="CLI for OnlyOffice Desktop Editors + RDF Knowledge Graphs",
    author="SLOANE OS",
    author_email="sloane@local",
    packages=find_packages(),
    python_requires=">=3.9",
    entry_points={
        "console_scripts": [
            "cli-anything-onlyoffice=cli_anything.onlyoffice.core.cli:main",
        ],
    },
    install_requires=[
        "python-docx>=1.1.0",
        "openpyxl>=3.1.2",
        "python-pptx>=0.6.23",
        "requests>=2.31.0",
        "scipy>=1.11.0",
        "rdflib>=7.0.0",
        "lxml>=4.9.0",
        "PyMuPDF>=1.24.0",
        "Pillow>=10.0.0",
        "pyshacl>=0.25.0",
    ],
    extras_require={
        "shacl": ["pyshacl>=0.25.0"],
    },
)
