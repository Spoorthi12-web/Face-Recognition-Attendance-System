# encode_faces.py — Reads all photos and saves face encodings

import os
import pickle
from deepface import DeepFace
from config import KNOWN_FACES_DIR

known_faces = []
known_names = []

print("Starting face encoding...")

# Loop through every photo in known_faces folder
for filename in os.listdir(KNOWN_FACES_DIR):
    if filename.endswith(".jpg") or filename.endswith(".png"):

        # Get person name from filename
        # example: john_doe.jpg → John Doe
        name = os.path.splitext(filename)[0]
        name = name.replace("_", " ").title()

        # Full path to the photo
        path = os.path.join(KNOWN_FACES_DIR, filename)

        print(f"Encoding: {name}...")

        # Generate face embedding using DeepFace
        embedding = DeepFace.represent(
            img_path   = path,
            model_name = "Facenet",
            enforce_detection = False
        )

        known_faces.append(embedding[0]["embedding"])
        known_names.append(name)
        print(f"Done: {name}")

# Save all encodings to file
with open("encodings.pkl", "wb") as f:
    pickle.dump({"encodings": known_faces, "names": known_names}, f)

print("---------------------------")
print(f"Total faces encoded: {len(known_names)}")
print("Saved to encodings.pkl")
print("Ready to run attendance!")
