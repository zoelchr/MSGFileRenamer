from setuptools import setup, find_packages

setup(
    name='msg_file_renamer',
    version='1.0',
    packages=find_packages(),
    url='',
    license='MIT License',
    author='Rüdiger Zölch',
    author_email='ruediger@zoelch.me',
    description='A utility for renaming MSG files efficiently based on their metadata.',
    package_data={
        '*': [
            'known_senders.csv',  # Include the CSV with list of known senders
            'msg_file_renamer.bat',  # Include the BAT for starting the utility
            'readme.md',   # Include the readme
            'testing.md',  # Include the testing directory
            'Das ist eine Test-Email.msg',  # Include a test email
        ]
    },
    include_package_data=True  # Ensure that package data is included
)