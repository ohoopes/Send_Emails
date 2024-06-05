from setuptools import setup, find_packages

setup(
    name='Send_Emails',
    version='0.1',
    packages=find_packages(),
    install_requires=[
        'requests',
        'msal',
        'pandas',
    ],
)
