from setuptools import setup, find_packages

setup(
    name='send_emails',  # Use lowercase and underscores for consistency
    version='0.1',
    packages=find_packages(),
    install_requires=[
        'requests',
        'msal',
        'pandas',
    ],
)
