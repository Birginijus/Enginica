from setuptools import setup


setup(
    name="excel-mysql-compare",
    version="0.1.0",
    description="Compare Excel worksheet rows with a MySQL table",
    long_description=open("README.md", encoding="utf-8").read(),
    long_description_content_type="text/markdown",
    py_modules=["excel_mysql_compare"],
    python_requires=">=3.9",
    install_requires=["openpyxl>=3.1", "PyMySQL>=1.1"],
    entry_points={"console_scripts": ["excel-mysql-compare=excel_mysql_compare:main"]},
)
