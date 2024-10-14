import cv2
from PyQt6.QtGui import QImage, QPixmap
from PyQt6.QtCore import QTimer, Qt

class CameraHandler:
    def __init__(self, camera_label):
        self.camera_label = camera_label
        self.capture = None
        self.timer = QTimer()
        self.timer.timeout.connect(self.update_frame)

    def start_camera(self, cam_index):
        if self.capture and self.capture.isOpened():
            self.capture.release()

        self.capture = cv2.VideoCapture(cam_index)
        if not self.capture.isOpened():
            print(f"Error: Unable to open camera {cam_index}")
            return

        self.capture.set(cv2.CAP_PROP_FRAME_WIDTH, 1280)
        self.capture.set(cv2.CAP_PROP_FRAME_HEIGHT, 720)
        self.timer.start(30)

    def update_frame(self):
        if self.capture and self.capture.isOpened():
            ret, frame = self.capture.read()
            if ret:
                frame = cv2.flip(frame, 1)
                image = cv2.cvtColor(frame, cv2.COLOR_BGR2RGB)
                height, width, channel = image.shape
                step = channel * width
                q_image = QImage(image.data, width, height, step, QImage.Format.Format_RGB888)
                self.camera_label.setPixmap(QPixmap.fromImage(q_image).scaled(
                    self.camera_label.size(), Qt.AspectRatioMode.KeepAspectRatio, Qt.TransformationMode.SmoothTransformation))
