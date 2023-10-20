import cv2
import openpyxl
import os
import time
import shutil
from tkinter import messagebox


class FaceRecognition:
    def __init__(self, face_detect_path, excel_path):
        self.video = cv2.VideoCapture(0)
        self.path = os.path.dirname(__file__)
        self.facedetect = cv2.CascadeClassifier(face_detect_path)
        self.excel_path = excel_path
        self.recognizer = cv2.face_LBPHFaceRecognizer.create()
        self.trainer_status = True
        trainer_file = os.path.join(self.path, "trainer.yml")
        try:
            self.recognizer.read(trainer_file)
        except:
            print(
                "Cảnh báo: Tệp 'trainer.yml' không tồn tại, trống hoặc không hợp lệ. Hãy đảm bảo bạn đã đào tạo mô hình trước."
            )
            self.trainer_status = False
        self.student_ids = self.read_ids_from_excel()

    def is_file_empty(file_path):
        return os.path.getsize(file_path) == 0

    def read_ids_from_excel(self):
        student_ids = {}
        try:
            workbook = openpyxl.load_workbook(os.path.join(self.path, "data1.xlsx"))
            sheet = workbook.active
            for row in sheet.iter_rows(min_row=2, values_only=True):
                print(row)
                if len(row) >= 3:
                    id, name, student_id = int(row[0]), row[1], row[2]
                    student_ids[id] = student_id

        except Exception as e:
            print(f"An error occurred while reading the Excel file: {e}")

        return student_ids

    def run(self):
        while True:
            ret, frame = self.video.read()
            if not ret:
                print("Không thể đọc khung hình từ camera.")
                break
            gray = cv2.cvtColor(frame, cv2.COLOR_BGR2GRAY)
            faces = self.facedetect.detectMultiScale(gray, 1.3, 5)

            for x, y, w, h in faces:
                serial, confident = self.recognizer.predict(gray[y : y + h, x : x + w])

                if confident < 50:
                    id = self.student_ids.get(serial)

                    cv2.putText(
                        frame,
                        f"{str(id).upper()}",
                        (x - 100, y + h - 300),
                        cv2.FONT_HERSHEY_COMPLEX,
                        1,
                        (0, 255, 0),
                        2,
                    )
                    cv2.rectangle(frame, (x, y), (x + w, y + h), (0, 255, 0), 2)

                    self.update_status_in_excel(serial)

                else:
                    cv2.putText(
                        frame,
                        "Unknown",
                        (x, y + h - 300),
                        cv2.FONT_HERSHEY_COMPLEX,
                        1,
                        (0, 0, 255),
                        2,
                    )
                    cv2.rectangle(frame, (x, y), (x + w, y + h), (0, 0, 255), 2)
                cv2.imshow("Face Recognition", frame)
            if cv2.waitKey(1) & 0xFF == ord("q"):
                break
        self.video.release()
        cv2.destroyAllWindows()

    def update_status_in_excel(self, student_id):
        self.create_backup()

        wb = openpyxl.load_workbook(self.excel_path)
        ws = wb.active
        found = False
        for row in ws.iter_rows(min_row=2, values_only=True):
            if row[0] == student_id:
                found = True
                row_index = ws._current_row
                if row[-1] == "V":
                    ws.cell(row=row_index, column=ws.max_column, value="X")
        if found:
            wb.save(self.excel_path)
        else:
            print(f"Không tìm thấy học sinh với student_id: {student_id}")

    def create_backup(self):
        if os.path.exists(self.excel_path):
            backup_file = os.path.join(os.path.dirname(__file__), "backup_data1.xlsx")

            shutil.copyfile(self.excel_path, backup_file)
