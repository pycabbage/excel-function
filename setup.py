from setuptools import setup

try:
    long_description = open("README.rst").read()
except IOError:
    long_description = ""

setup(
    name="excel_func",
    version="0.0.0",
    license="MIT",
    author="pycabbage",
    packages=["excel_func"],
    package_dir={"excel_func": "src"},
    description="excel_func",
    long_description=long_description,
    install_requires=open("requirements.txt").read().strip().splitlines(),
    entry_points={
        "console_scripts": [
            "excel_func = excel_func.__main__:console_wrapper",
        ]
    }
)
