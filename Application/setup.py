from setuptools import setup

setup(
    name="Application",
    version="1.0",
    license="MIT",
    packages=["Application.App"],
    include_package_data=True,
    install_requires=[
      "python-pptx","openpyxl","PyQt5"
    ],
)