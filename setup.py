from setuptools import find_packages,setup

setup(
    name='Barcode-Scanner',
    version='0.0.1',
    author='suvradeep',
    author_email='dassuvradeep9@gmail.com',
    install_requires=["opencv-contrib-python==4.7.0.72","opencv-python==4.8.0.74","opencv-python-headless==4.8.0.74","pythoncom","pyzbar","win32com.client","tkinter","pandas","openpyx"],
    packages=find_packages()
)