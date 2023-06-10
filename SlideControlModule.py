import keyboard
import pyautogui

class SlideControl:
    def __init__(self, slide_change_threshold=80):
        self.slide_change_threshold = slide_change_threshold

    def detect_slide_change(self, length):
        if length < self.slide_change_threshold:
            # Use keyboard shortcuts to change slides
            pyautogui.press('right')  # Simulate the "right" arrow key to move to the next slide
        else:
            x1, y1 = lmList[4][1], lmList[4][2]
            x2, y2 = lmList[8][1], lmList[8][2]
            if abs(x2 - x1) > self.slide_change_threshold:
                # Move to the next or previous slide based on hand movement direction
                if x2 > x1:
                    # Move to the next slide
                    slide_number = presentation.SlideShowWindow.View.Slide.SlideIndex
                    if slide_number < slide_count:
                        presentation.SlideShowWindow.View.Next()
                    else:
                        presentation.SlideShowWindow.View.First()
                else:
                    # Move to the previous slide
                    slide_number = presentation.SlideShowWindow.View.Slide.SlideIndex
                    if slide_number > 1:
                        presentation.SlideShowWindow.View.Previous()
                    else:
                        presentation.SlideShowWindow.View.Last()
