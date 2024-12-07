from setuptools import setup, find_packages

setup(
    name="async-excel",
    version="0.1.0",
    author="a123ao",
    description="async-excel is a library for reading and writing Excel files in async mode.",
    long_description=open("README.md").read(),
    long_description_content_type="text/markdown",
    url="https://github.com/a123ao/async-excel",
    package_dir={"": "src"},
    packages=find_packages(where="src"),
    classifiers=[
        "Programming Language :: Python :: 3.8",
        "License :: OSI Approved :: MIT License",
        "Operating System :: Microsoft :: Windows",
        "Development Status :: 3 - Alpha",
        "Intended Audience :: Developers",
        "Topic :: Software Development :: Libraries :: Python Modules",
    ],
    requires=["pywin32"],
)