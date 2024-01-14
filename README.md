# Attendance System Using Code 39 Barcode

This Project about Attendance System is utilizing Barcode Scanning method is a Python project tailored for college students and professors (Professors dataset not available hence thispart is Pending). Employing tkinter and pyzbar libraries, it captures barcode information from ID cards through a camera and fetches corresponding details from an Excel sheet. This automated system efficiently updates attendance records in Excel, streamlining the process and minimizing errors. Customized for Core Courses, it provides a user-friendly solution for effective attendance management in educational institutions.

#  Primary contributions
1. Automated Attendance Management System: Implements automated attendance tracking to save time and minimize errors.
2. Barcode Scanning Technology: Effectively monitors attendance through distinctive barcodes.
3. Real-time Feedback Mechanism: Delivers immediate attendance updates to enhance transparency.
4. User-friendly Interface: Simplifies interaction for both students and faculty members.
5. Security Enhancement: Guarantees system access exclusively for authorized individuals.

# Note
1. This architecture utilizes Code 39 as our college uses Code 39, Code 39 is cost effective and straightforward to implement. It doesn't  demand intricate encoding or specialized equipment.
2. I had no data related to faculty list hence The faculty attendance part is pending

# Steps and Architecture
1. Module Integration: To ensure the seamless operation of our system, I incorporated various essential modules such as pywin32, opencv, time, pyzbar, tkinter, datetime, and openpyxl. These modules serve distinct functionalities to system, contributing to its smooth execution.

2. Excel File and Sheet Management: The system features capabilities that enable users to access attendance records, course and faculty information, and student records in a secure view. This functionality includes concealing sheets that are not currently in use and safeguarding sheets with a password. Additionally, users have the option to input a selected date into a specific cell in an Excel sheet for added convenience.

3. Barcode Scanning: The central functionality of the system revolves around barcode scanning. We utilize the OpenCV library to continuously capture frames from a webcam and identify barcodes within those frames. Upon detecting a new barcode, it is added to the list of barcodes. Subsequently, we employ the pyzbar library to decode the barcodes identified by the system.

4. User Authentication and Attendance Logging: In this system, I prioritize user authentication and precise attendance logging. We have implemented the following functionalities:

User Attendance Verification: The system scans a barcode, extracts data from an Excel sheet, and confirms the user's attendance by cross-referencing their roll number with the data in the sheet.

Faculty Attendance Marking (Pending due to the lack of data)

5. Employing SVM for Barcode Categorization: In order to classify barcode information and determine the attendance status of students, we utilize Support Vector Machine (SVM) algorithms. SVMs prove highly effective in accurately categorizing different types of barcodes. Preprocessing techniques, implemented using OpenCV, are applied to enhance classification accuracy for barcode images.

6. Excel Database for Attendance Administration: This Architecture relies on an Excel workbook as the primary database for storing diverse datasets. These datasets encompass student data, course information, and attendance records. Student data includes details such as names, roll numbers, and other pertinent information. Course data provides details about the courses being taught, while attendance data captures information on student attendance, including the date and attendance status.

7. Image Processing Strategies: Image processing plays a pivotal role in our system, particularly for barcode extraction and decoding. The process initiates with image acquisition, where a camera captures the barcode image. Techniques such as thresholding, smoothing, edge detection, and contour analysis are employed to enhance barcode quality, eliminate noise, and delineate barcode boundaries. Barcode decoding relies on specialized software or libraries capable of recognizing and decoding various barcode types.

# Summary
This module encompasses crucial code snippets necessary for the seamless operation of the overall project. The decision to create a distinct module file enhances code management and organization, fostering a more modular and maintainable codebase. The project itself constitutes a Graphical User Interface (GUI) fashioned using tkinter, facilitating students and faculty members from the Cyber Security Batch at my university to execute diverse functions. These functions include adding, viewing, and modifying attendance, as well as accessing vital databases. To decode barcodes, the project leverages the pyzbar library and implements an SVM Classifier for precise and efficient recognition. Additionally, OpenCV is employed to capture barcodes and process corresponding images. These tools and techniques collectively enable the system to swiftly and accurately extract information from barcodes, matching it with the relevant student or faculty member. For database management, the project adopts Excel files, a widely used and easily accessible format for the majority of users. These Excel files store crucial data, encompassing student and faculty information, along with attendance records. By utilizing Excel files, the project offers a straightforward yet effective solution for database management, simplifying the process for users to modify and access data as needed.

# Structure of each File in this Repo

1. `main.py` - This script serves as the primary Python file housing all essential functionalities and modules for the project. It functions as the entry point for the Graphical User Interface (GUI), enabling users to perform a range of actions, including adding, viewing, and modifying attendance, accessing crucial databases, and decoding barcodes. The "attendance_system.py" file stands as the backbone of the project, playing a pivotal role in ensuring the proper functioning of the system.

2. `module.py` - Developed to host various helper functions and utilities utilized by the main file, this module file is imported and employed by the main script for tasks such as image processing, barcode decoding, and database management.

3. `decoder.py` - This code segment, residing within the Module.py file, captures and decodes barcodes using the Pyzbar library. Leveraging the OpenCV library, the function captures frames from the camera, decoding any barcodes present in those frames using Pyzbar. The decoded barcodes are then stored in a list, providing the program with essential information.

4. `screen.py` - Nested within the "Module.py" file, this code snippet is responsible for generating a loading screen featuring a progress bar while the application loads. Enhancing user experience, the loading screen offers a clear indication that the application is loading, minimizing the perception of waiting time.

5. `Excel.py` - This code excerpt, located in the "module.py" file, facilitates the viewing of specific Excel sheets in a protected view. Its purpose is to enable users to view specific sheets in a workbook without granting them the ability to modify or make changes to the content. This is achieved by concealing all other sheets in the workbook and activating a protected view that restricts the user's ability to edit or interact with the content.

6. `requirement.txt` - This module is to import all required libraries. The libraries used in this project are OpenCV, tkinter, datetime, openpyxl, pyzbar, time, datetime, pywin32, pythoncom
