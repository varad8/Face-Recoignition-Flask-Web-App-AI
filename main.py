import pyrebase
from flask import Flask, flash, redirect, render_template, request, session, abort, url_for,Response,json
import cv2
import numpy as np
import face_recognition
import os
import time
import json
from win32com.client import Dispatch
import threading
import datetime



def speak(str1):
    speak=Dispatch(("SAPI.SpVoice"))
    speak.Speak(str1)


app = Flask(__name__)       #Initialze flask constructor

camera = cv2.VideoCapture(0)




#Add your own details
config = {
  "apiKey": "AIzaSyB5KBwePYXsOXo6lsJXl0fkW_-OLzmvUrY",
  "authDomain": "face-recognization-6879a.firebaseapp.com",
  "databaseURL": "https://face-recognization-6879a-default-rtdb.firebaseio.com/",
  "storageBucket": "face-recognization-6879a.appspot.com"
}

#initialize firebase
firebase = pyrebase.initialize_app(config)
auth = firebase.auth()
db = firebase.database()

# Dictionary to keep track of attendance status for each student
attendance_status = {}



# Login
@app.route("/")
def login():
    return render_template("login.html")


# Log out page
@app.route("/logout", methods=["GET"])
@app.route("/logout")
def logout():
    session.clear()
    return redirect(url_for('login'))

@app.route("/dashboard")
def dashboard():
    try:
        if "is_logged_in" in session:
            # Fetch all attendance data
            all_attendance_data = db.child("attendance").get().val()

            attendance_record=all_attendance_data

            # Check if all_students is None
            if attendance_record is None:
                attendance_record = {}


            # Fetch all student records
            all_students = db.child("students").get().val()
            
            # Check if all_students is None
            if all_students is None:
                all_students = {}

            print(f"Debug - Raw Attendance Data: {all_attendance_data}")

            # Check if all_attendance_data is None
            if all_attendance_data is None:
                all_attendance_data = {}

            print(f"Debug - Processed Attendance Data: {all_attendance_data}")

            # Fetch other attendance data
            yearly_counts = {}
            todays_count = {}
            this_month_count = {}

            for month_year, student_data in all_attendance_data.items():
                year = month_year.split('_')[1]
                if year not in yearly_counts:
                    yearly_counts[year] = 0

                for roll_no, attendance_data in student_data.items():
                    for attendance_date, attendance_info in attendance_data.items():
                        # Check if the year is the current year
                        if year == str(datetime.datetime.now().year):
                            yearly_counts[year] += 1

                        # Check for today's count
                        today_date = datetime.date.today().strftime("%Y-%m-%d")
                        if today_date == attendance_date and attendance_info.get("present", False):
                            todays_count[today_date] = todays_count.get(today_date, 0) + 1

                        # Check for this month's count
                        current_month = datetime.datetime.now().month
                        attendance_month = int(attendance_date.split('-')[1])
                        if current_month == attendance_month and attendance_info.get("present", False):
                            month_name = datetime.date(1900, current_month, 1).strftime('%B')
                            this_month_count[month_name] = this_month_count.get(month_name, 0) + 1

            # Count of all student records
            total_students_count = len(all_students)

            print("Debug - Yearly Counts:", yearly_counts)
            print("Debug - Todays Count:", todays_count)
            print("Debug - This Month's Count:", this_month_count)

      


            return render_template("dashboard.html",email = session["email"],name = session["name"] ,
                                   yearly_counts=yearly_counts, todays_count=todays_count,
                                   this_month_count=this_month_count,total_students_count=total_students_count,attendance_record=attendance_record)
        else:
            return redirect(url_for('login'))
    except Exception as e:
        print(f"Error in dashboard route: {e}")
        return render_template("error.html")
       


# Get All Presenty For that Student
def get_attendance_data(student_roll_no):
    # Fetch all attendance data
    all_attendance_data = db.child("attendance").get().val()

    # Initialize a list to store filtered attendance data
    filtered_attendance_data = []

    # Check if attendance data is not empty
    if all_attendance_data:
        # Iterate over the months and their respective attendance data
        for month_year, month_data in all_attendance_data.items():
            # Check if the student roll number exists in the current month's data
            if student_roll_no in month_data:
                # Append the relevant month's data to the filtered list
                filtered_attendance_data.append((month_year, {student_roll_no: month_data[student_roll_no]}))

    return filtered_attendance_data

@app.route("/student_list")
def student_list():
    # Check if the user is logged in
    if not "is_logged_in" in session:
        return redirect(url_for('login'))

    # Fetch all students data from Firebase
    all_students_data = db.child("students").get()

    if all_students_data.each():
        all_students = []
        for student_key, student_data in all_students_data.val().items():
            roll_no = student_data.get("roll_no")
            folder_path = f"static/images/{roll_no}"

            # List all files in the folder
            image_files = os.listdir(folder_path)

            # Sort the list to ensure a consistent order (you may want to sort differently)
            image_files.sort()

            # Check if there are images in the folder
            if image_files:
                # Get the path of the first image (position 0)
                image_path = os.path.join(folder_path, image_files[0])
            else:
                image_path = None

            all_students.append({
                "name": student_data['name'],
                "roll_no": roll_no,
                "dob": student_data.get("dob"),
                "address": student_data.get("address"),
                "mobile_no": student_data.get("mobile_no"),
                "image_path": image_path,
            })

        selected_student = None
        image_path = None

        # Check if a student is selected (via query parameter)
        selected_roll_no = request.args.get("roll_no")
        if selected_roll_no:
            # Fetch the selected student's data from Firebase
            selected_student_data = db.child("students").order_by_child("roll_no").equal_to(selected_roll_no).get()

            if selected_student_data.each():
                # Assuming there is only one student with the given roll number
                selected_student = list(selected_student_data.val().values())[0]

                # Path to the folder containing selected student images
                folder_path = f"static/images/{selected_roll_no}"

                # List all files in the folder
                image_files = os.listdir(folder_path)

                # Sort the list to ensure a consistent order (you may want to sort differently)
                image_files.sort()

                # Check if there are images in the folder
                if image_files:
                    # Get the path of the first image (position 0)
                    image_path = os.path.join(folder_path, image_files[0])

        # Fetch attendance data for the selected student
        attendance_data = get_attendance_data(selected_roll_no)

        return render_template("student_list.html", all_students=all_students, selected_student=selected_student, image_path=image_path, attendance_data=attendance_data)

    else:
        flash("No students data found in the database.", "error")
        return redirect(url_for('student_list'))



# If someone clicks on login, they are redirected to /result
@app.route("/result", methods=["POST", "GET"])
def result():
    if request.method == "POST":
        result = request.form
        email = result["email"]
        password = result["pass"]
        try:
            user = auth.sign_in_with_email_and_password(email, password)
            user_data = db.child("admin").get().val().get(user["localId"])

            if user_data:
                session['is_logged_in'] = True
                session['email'] = user["email"]
                session['uid'] = user["localId"]
                session['name'] = user_data["name"]

                # Redirect to the dashboard page instead of the welcome page
                return redirect(url_for('dashboard'))
            else:
                error_message = "User data not found. Please check your credentials."
                return render_template('login.html', error=error_message)
        except Exception as e:
            error_message = "Login failed. Please check your credentials."
            return render_template('login.html', error=error_message)
    else:
        if session.get('is_logged_in'):
            return redirect(url_for('dashboard'))
        else:
            return redirect(url_for('login'))
            



# Add Student route with form and webcam capture
@app.route("/add_student", methods=["GET", "POST"])
def add_student():
    # Check if the user is logged in
    if not "is_logged_in" in session:
        return redirect(url_for('login'))

    if request.method == "POST":
        try:
            # Handle form submission and save student details
            # You may need to process the form data, validate, and save it to the database
            # For simplicity, let's assume you have a form with fields: name, roll_no, dob, address, mobile_no

            name = request.form.get("name")
            roll_no = request.form.get("roll_no")
            dob = request.form.get("dob")
            address = request.form.get("address")
            mobile_no = request.form.get("mobile_no")

        
            # Capture 10 face images
            capture_images(roll_no,name)

            # Check if the roll number already exists in the database
            existing_student = db.child("students").order_by_child("roll_no").equal_to(roll_no).get()

            if existing_student.each():
                speak("Student with this roll number already exists.")
            else:
                # Save the student details to the database
                db.child("students").push({
                    "name": name,
                    "roll_no": roll_no,
                    "dob": dob,
                    "address": address,
                    "mobile_no": mobile_no,
                    "password": "user_" + roll_no + "$#&CBT"
                })
                
                speak("Images Captured and saved data")
             
           

            flash("Student added successfully.", "success")
            return render_template("add_student.html")

        except Exception as e:
            flash(f"An error occurred: {str(e)}", "error")
            return render_template("add_student.html")

    return render_template("add_student.html")


# Function to capture 10 face images and save them
def capture_images(roll_no,name):
    try:
        speak("Capturing Images")

        # Create a folder for the student using their roll number
        folder_path = f"static/images/{roll_no}"
        os.makedirs(folder_path, exist_ok=True)

        # Capture 10 face images
        for i in range(10):
            success, frame = camera.read()
            if not success:
                break

            # Perform face detection
            faces = detect_faces(frame)

            if len(faces) > 0:
                # Get the first detected face
                x, y, w, h = faces[0]

                # Crop the face region
                face_image = frame[y:y + h, x:x + w]

                # Save the face image
                image_path = os.path.join(folder_path, f"{name}_{roll_no}_{i + 1}.jpg")
                cv2.imwrite(image_path, face_image)
                speak(f"Image {i + 1} captured for roll number {roll_no}")
              

        # Release the camera
        # camera.release()

    except Exception as e:
        print(f"An error occurred while capturing images: {str(e)}")



# Function to detect faces in a given frame
def detect_faces(frame):
    detector = cv2.CascadeClassifier(cv2.data.haarcascades + 'haarcascade_frontalface_default.xml')
    faces = detector.detectMultiScale(frame, 1.1, 7)
    return faces



#Generatiing frames for adding faces
def gen_frames():  
    while True:
        success, frame = camera.read()  # read the camera frame
        if not success:
            break
        else:
            detector=cv2.CascadeClassifier(cv2.data.haarcascades + 'haarcascade_frontalface_default.xml')
            faces=detector.detectMultiScale(frame,1.1,7)
             #Draw the rectangle around each face
            for (x, y, w, h) in faces:
                cv2.rectangle(frame, (x, y), (x+w, y+h), (255, 0, 0), 2)

            
            ret, buffer = cv2.imencode('.jpg', frame)
            frame = buffer.tobytes()
            yield (b'--frame\r\n'
                   b'Content-Type: image/jpeg\r\n\r\n' + frame + b'\r\n')

#Getting Live Preview of Web Cam
@app.route('/video_feed')
def video_feed():
    return Response(gen_frames(), mimetype='multipart/x-mixed-replace; boundary=frame')


# Mark Attendance 
def mark_attendance(student_roll_no):
    # Get current month and year
    now = datetime.datetime.now()
    month_year = f"{now.month}_{now.year}"
    current_time = now.strftime("%I:%M:%S %p")

    # Check if the attendance collection for the current month exists
    if not db.child("attendance").child(month_year).get().val():
        # Create the month-wise collection
        db.child("attendance").child(month_year).set({})

    # Check if the student has already been marked present today
    today = datetime.date.today().isoformat()
    attendance_data = db.child("attendance").child(month_year).child(student_roll_no).child(today).get().val()

    if attendance_data is None:
        # Mark the student's attendance for the current day with timestamp
        attendance_info = {
            "present": True,
            "timestamp": current_time
        }
        db.child("attendance").child(month_year).child(student_roll_no).child(today).set(attendance_info)
    else:
        print(f"Attendance already marked for {today} for student {student_roll_no}")


# Shared variable for storing camera frames
output_frame = None
lock = threading.Lock()

def generate():
    global output_frame
    while True:
        with lock:
            if output_frame is not None:
                ret, buffer = cv2.imencode('.jpg', output_frame)
                frame_bytes = buffer.tobytes()
                yield (b'--frame\r\n'
                       b'Content-Type: image/jpeg\r\n\r\n' + frame_bytes + b'\r\n')



def detect_live_faces(camera, all_students):
    global output_frame
    global lock

    face_cascade = cv2.CascadeClassifier(cv2.data.haarcascades + 'haarcascade_frontalface_default.xml')

    while True:
        success, frame = camera.read()  # read the camera frame
        if not success:
            break

        gray = cv2.cvtColor(frame, cv2.COLOR_BGR2GRAY)

        # Face detection
        faces = face_cascade.detectMultiScale(gray, scaleFactor=1.1, minNeighbors=5, minSize=(30, 30))

        with lock:
            for (x, y, w, h) in faces:
                # Draw rectangle and label on the face
                cv2.rectangle(frame, (x, y), (x+w, y+h), (0, 255, 0), 2)

                # Check against stored faces in the Firebase student data
                for student in all_students:
                    roll_no = student.get("roll_no")
                    student_folder = f"static/images/{roll_no}"

                    # Check if the student folder exists
                    if os.path.exists(student_folder):
                        # Iterate over the images in the student folder
                        for file_name in os.listdir(student_folder):
                            if file_name.endswith('.jpg'):
                                stored_face_image = cv2.imread(os.path.join(student_folder, file_name), cv2.IMREAD_GRAYSCALE)
                                stored_face_resized = cv2.resize(stored_face_image, (w, h))

                                # Compare the detected face with the stored face
                                result = cv2.matchTemplate(gray[y:y+h, x:x+w], stored_face_resized, cv2.TM_CCOEFF_NORMED)
                                _, confidence, _, _ = cv2.minMaxLoc(result)

                                if confidence > 0.7:  # You can adjust this threshold
                                    # Draw a box around the face
                                    cv2.rectangle(frame, (x, y), (x+w, y+h), (0, 0, 255), 2)

                                    # Get the size of the text
                                    label_size, baseline = cv2.getTextSize(student['name'], cv2.FONT_HERSHEY_SIMPLEX, 0.5, 1)
                                    
                                    # Draw a rectangle for the label dynamically based on text size
                                    cv2.rectangle(frame, (x, y + h), (x + label_size[0] + 6, y + h + label_size[1] + 6), (0, 0, 255), cv2.FILLED)
                                    
                                    # Draw a label with the student's name below the face
                                    font = cv2.FONT_HERSHEY_SIMPLEX
                                    cv2.putText(frame, student['name'], (x + 3, y + h + label_size[1] + 3),
                                                font, 0.5, (255, 255, 255), 1)
                                    
                                    
                                    # Mark the student as present for today
                                    if roll_no not in attendance_status:
                                        attendance_status[roll_no] = {"marked_today": True, "date": datetime.date.today()}
                                        mark_attendance(roll_no)


            output_frame = frame

def start_face_detection(camera, all_students):
    threading.Thread(target=detect_live_faces, args=(camera, all_students)).start()

@app.route("/check_faces_match")
def check_faces_match():
    # Fetch all students data from Firebase
    all_students_data = db.child("students").get()

    if all_students_data.each():
        all_students = []
        for student_key, student_data in all_students_data.val().items():
            all_students.append({
                "name": student_data['name'],
                "roll_no": student_data.get("roll_no"),
              
            })

        camera = cv2.VideoCapture(0)  # Assuming camera index 0, change it if needed

        # Start the face detection thread
        start_face_detection(camera, all_students)

        return Response(generate(), mimetype='multipart/x-mixed-replace; boundary=frame')

    else:
        flash("No students data found in the database.", "error")
        return redirect(url_for('check_faces_match'))


    
############################# Student Route   ###################################################################################

# Student Login route
@app.route("/student_login")
def student_login():
    return render_template("student_login.html", error=None)


# Student Result route
@app.route("/student_result", methods=["POST"])
def student_result():
    if request.method == "POST":
        roll_no = request.form.get("roll_no")
        password = request.form.get("password")

        # Retrieve student data from the database based on the provided roll number
        student_data = db.child("students").order_by_child("roll_no").equal_to(roll_no).get()

        if student_data.each():
            # Assuming there is only one student with the given roll number
            student = list(student_data.val().values())[0]

            # Check if the provided password matches the stored password
            if password == student["password"]:
                # Set the student information in the session or any other desired storage
                session["student_roll_no"] = roll_no
                session["student_name"] = student["name"]

                # Redirect to the student dashboard or any other desired page
                return redirect(url_for('student_dashboard'))
            else:
                error_message = "Incorrect password. Please try again."
        else:
            error_message = "Student with this roll number does not exist."

        # Flash an error message and render the login template with the error
        flash(error_message, "error")
        return render_template("student_login.html", error=error_message)



# Student Dashboard page
@app.route("/student_dashboard")
def student_dashboard():
    # Check if the student is logged in
    if "student_roll_no" in session:
        roll_no = session["student_roll_no"]
        name = session["student_name"]

        #Get Presenty Data
        attendance_data=get_attendance_data(roll_no)
        current_date = datetime.date.today().isoformat()
     

        # Path to the folder containing student images
        folder_path = f"static/images/{roll_no}"

        # List all files in the folder
        image_files = os.listdir(folder_path)

        # Sort the list to ensure a consistent order (you may want to sort differently)
        image_files.sort()

        # Check if there are images in the folder
        if image_files:
            # Get the path of the first image (position 0)
            image_path = os.path.join(folder_path, image_files[0])

            # You can pass this image_path to your HTML template and display it using an <img> tag
            return render_template("student_dashboard.html", roll_no=roll_no, name=name, image_path=image_path,attendance_data=attendance_data,current_date=current_date)
        else:
            # Handle the case when there are no images in the folder
            return render_template("student_dashboard.html", roll_no=roll_no, name=name, image_path=None)
    else:
        # Redirect to the login page if the student is not logged in
        return redirect(url_for('student_login'))


if __name__ == "__main__":
    app.secret_key = "varad123456"
    app.run(debug=True)