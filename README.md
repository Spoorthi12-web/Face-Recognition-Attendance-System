# Face Recognition Attendance System

AI-powered attendance system using Face Recognition, Python, DeepFace, MySQL and Flask.

## Features
- Real-time face recognition using DeepFace and OpenCV
- Punch In / Punch Out with live webcam
- MySQL database with 3 tables
- Beautiful Excel attendance reports
- Flask web dashboard

## Tech Stack
- Python 3.13
- DeepFace + OpenCV
- MySQL
- Flask
- pandas + openpyxl

## Setup
1. Install dependencies: `pip install deepface opencv-python pandas openpyxl flask mysql-connector-python`
2. Configure `config.py` with your MySQL credentials
3. Add photos to `known_faces/` folder
4. Run `encode_faces.py` to register faces
5. Run `attendance.py` to start the system
6. Run `app.py` for web dashboard

## Developer
Built by Spoorthi — Engineering Student

GitHub: [Spoorthi12-web](https://github.com/Spoorthi12-web)

