# Face Recognition Attendance System

AI-powered attendance system using Face Recognition, Python, DeepFace, MySQL, Flask and Gmail automation.

## Features
- Real-time face recognition using DeepFace and OpenCV
- Punch In / Punch Out with live webcam
- MySQL database with 3 tables
- Beautiful Excel attendance reports
- Flask web dashboard
- Automated daily and weekly email reports to manager

## Tech Stack
- Python 3.13
- DeepFace + OpenCV
- MySQL
- Flask
- pandas + openpyxl
- smtplib (Gmail automation)

## Project Structure
- `encode_faces.py` — Register known faces
- `attendance.py` — Live webcam detection + Punch In/Out
- `report_generator.py` — Generate Excel reports
- `email_sender.py` — Automated Gmail reports
- `app.py` — Flask web dashboard

## Setup
1. Install dependencies: `pip install deepface opencv-python pandas openpyxl flask mysql-connector-python`
2. Configure `config.py` with your MySQL and Gmail credentials
3. Add photos to `known_faces/` folder
4. Run `encode_faces.py` to register faces
5. Run `attendance.py` to start the system
6. Run `app.py` for web dashboard

## Screenshots
Coming soon!

## Developer
Built by Spoorthi — Engineering Student
```

