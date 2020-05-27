from setuptools import setup, find_packages

setup(
    name='prodctrlcore',
    version=0.1,
    packages=find_packages("src"),
    package_dir={"": "src"},
    install_requires=[
        "xlwings>=0.19.4",
        "pyodbc>=4.0.30",
        "pandas>=1.0.3",
        "graphqlclient>=0.2.4",
        "inflection>=0.4.0"
    ]
)
