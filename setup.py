import io
import re

from setuptools import find_packages
from setuptools import setup

with io.open("src/brm_cue_pdf/__init__.py", "rt", encoding="utf8") as f:
    version = re.search(r'__version__ = "(.*?)"', f.read()).group(1)

setup(
    name="brm_cue_pdf",
    version=version,
    packages=find_packages("src"),
    install_requires=[
        "reportlab",
        "xlrd"
    ],
    extras_require={
        "test": [
            "pytest",
        ],
    },
    package_dir={"": "src"},
    python_requires=">=3.6",
    entry_points={"console_scripts": [
        "brm_cue_pdf = brm_cue_pdf.cli:main",
    ]},
)
