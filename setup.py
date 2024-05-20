from setuptools import setup, find_packages

setup(
    name='excel2json',
    version='1.0.0',
    packages=find_packages(),
    install_requires=[
        'xlrd',
        'PyYAML',
    ],
    entry_points={
        'console_scripts': [
            'excel2json=excel_to_json.main:main',
        ],
    },
)
