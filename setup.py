from setuptools import setup, find_packages

setup(
    name='hss_core',
    version=1.0,
    packages=find_packages("src"),
    package_dir={"": "src"},
    install_requires=[
        "xlwings",
        "pyodbc",
        "pandas",
        "graphqlclient",
        "inflection"
    ]
)
