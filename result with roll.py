import cv2
import pandas as pd
import os

# === Paths ===
input_folder = r'C:\Users\Anand\Downloads\OMR\OMR\OMR'
answer_key_path = r"C:\Users\Anand\Downloads\OMR\OMR\ANSWERKEY\ANSWERKEY_0001.tif"
output_file = r'C:\omr\omr_results.xlsx'

# === Question Coordinates ===
coordinates = {
    1: {'A': (180, 956, 240, 1026), 'B': (180, 1020, 244, 1088), 'C': (182, 1088, 242, 1144), 'D': (176, 1146, 242, 1222)},
    2: {'A': (260, 956, 332, 1020), 'B': (260, 1022, 332, 1088), 'C': (262, 1086, 334, 1144), 'D': (262, 1148, 332, 1222)},
    3: {'A': (340, 958, 416, 1024), 'B': (342, 1022, 414, 1090), 'C': (340, 1088, 414, 1144), 'D': (340, 1146, 418, 1220)},
    4: {'A': (428, 958, 496, 1022), 'B': (428, 1024, 496, 1090), 'C': (428, 1088, 498, 1144), 'D': (428, 1148, 496, 1218)},
    5: {'A': (510, 960, 586, 1020), 'B': (510, 1020, 586, 1086), 'C': (510, 1088, 584, 1144), 'D': (512, 1146, 584, 1222)},
    6: {'A': (588, 956, 666, 1022), 'B': (590, 1024, 668, 1086), 'C': (590, 1086, 668, 1144), 'D': (590, 1146, 668, 1220)},
    7: {'A': (674, 960, 748, 1022), 'B': (674, 1024, 750, 1084), 'C': (674, 1086, 748, 1144), 'D': (674, 1148, 748, 1222)},
    8: {'A': (758, 960, 832, 1022), 'B': (758, 1022, 832, 1088), 'C': (758, 1086, 834, 1144), 'D': (758, 1146, 834, 1222)},
    9: {'A': (842, 960, 916, 1024), 'B': (842, 1024, 914, 1086), 'C': (842, 1086, 914, 1144), 'D': (842, 1146, 914, 1222)},
    10: {'A': (926, 958, 996, 1024), 'B': (926, 1024, 998, 1088), 'C': (926, 1086, 998, 1146), 'D': (926, 1146, 998, 1222)},
    11: {'A': (180, 1280, 248, 1344), 'B': (182, 1344, 250, 1408), 'C': (182, 1408, 248, 1472), 'D': (182, 1470, 248, 1540)},
    12: {'A': (262, 1282, 334, 1344), 'B': (264, 1346, 334, 1408), 'C': (264, 1410, 332, 1470), 'D': (262, 1470, 332, 1540)},
    13: {'A': (340, 1280, 418, 1346), 'B': (340, 1346, 416, 1406), 'C': (340, 1410, 418, 1466), 'D': (342, 1470, 416, 1542)},
    14: {'A': (428, 1282, 502, 1344), 'B': (428, 1344, 502, 1408), 'C': (428, 1410, 502, 1470), 'D': (428, 1470, 502, 1540)},
    15: {'A': (510, 1282, 586, 1342), 'B': (510, 1344, 584, 1408), 'C': (510, 1410, 584, 1472), 'D': (510, 1470, 586, 1540)},
    16: {'A': (588, 1282, 664, 1344), 'B': (590, 1346, 668, 1408), 'C': (590, 1410, 664, 1472), 'D': (590, 1470, 666, 1536)},
    17: {'A': (672, 1280, 748, 1346), 'B': (672, 1346, 748, 1408), 'C': (674, 1410, 748, 1472), 'D': (674, 1472, 748, 1540)},
    18: {'A': (758, 1280, 832, 1342), 'B': (758, 1344, 832, 1408), 'C': (758, 1410, 832, 1472), 'D': (758, 1470, 832, 1538)},
    19: {'A': (840, 1282, 914, 1344), 'B': (842, 1346, 914, 1410), 'C': (840, 1410, 912, 1472), 'D': (840, 1472, 914, 1542)},
    20: {'A': (898, 1282, 990, 1346), 'B': (902, 1346, 988, 1410), 'C': (900, 1410, 990, 1466), 'D': (900, 1466, 990, 1538)},
}

# === Roll Number Coordinates ===
roll_coords = {
    'F': [(198, 240, 276, 292), (200, 286, 278, 352), (196, 348, 276, 424), (200, 424, 278, 482), (198, 482, 280, 542),
          (198, 538, 280, 600), (200, 602, 282, 672), (198, 666, 280, 732), (200, 726, 280, 790), (196, 784, 282, 840)],
    'S': [(278, 240, 362, 292), (282, 292, 364, 354), (276, 352, 364, 422), (280, 424, 364, 480), (282, 480, 364, 546),
          (282, 550, 366, 600), (282, 606, 364, 664), (284, 670, 366, 722), (282, 726, 366, 792), (286, 786, 368, 846)],
    'T': [(364, 240, 454, 290), (360, 286, 450, 354), (366, 354, 448, 422), (366, 422, 452, 486), (370, 482, 452, 544),
          (368, 544, 452, 596), (368, 606, 454, 666), (368, 664, 454, 726), (372, 732, 454, 790), (368, 792, 454, 842)],
    'L': [(454, 234, 526, 292), (452, 292, 528, 360), (452, 356, 528, 418), (454, 418, 530, 480), (454, 486, 530, 546),
          (456, 544, 526, 606), (456, 606, 524, 670), (460, 670, 526, 734), (454, 734, 526, 790), (458, 796, 526, 848)],
}

def is_marked(region, white_threshold=220, pixel_ratio=0.02):
    gray = cv2.cvtColor(region, cv2.COLOR_BGR2GRAY)
    total_pixels = gray.size
    dark_mask = (gray < white_threshold).astype('uint8') * 255
    dark_pixels = cv2.countNonZero(dark_mask)
    return (dark_pixels / total_pixels) > pixel_ratio

def detect_roll_number(image):
    roll = ""
    for section in ['F', 'S', 'T', 'L']:
        digit_detected = None
        for i, (x1, y1, x2, y2) in enumerate(roll_coords[section]):
            roi = image[y1:y2, x1:x2]
            if is_marked(roi):
                if digit_detected is not None:
                    return "Invalid"
                digit_detected = str(i)
        if digit_detected is None:
            return "Invalid"
        roll += digit_detected
    return roll

def extract_answers_from_image(image):
    answers = {}
    for qno, options in coordinates.items():
        marked_options = []
        for opt, (x1, y1, x2, y2) in options.items():
            roi = image[y1:y2, x1:x2]
            if is_marked(roi):
                marked_options.append(opt)
        if len(marked_options) == 1:
            answers[qno] = marked_options[0]
        else:
            answers[qno] = "Invalid" if len(marked_options) > 1 else "Unmarked"
    return answers

# === Answer Key ===
ans_key_img = cv2.imread(answer_key_path)
answer_key = extract_answers_from_image(ans_key_img)
answer_key_row = {'Image': 'Answer Key', 'RollNumber': '---', 'Obtained Marks': '---'}
for q in range(1, 21):
    answer_key_row[f'Ans_{q}'] = answer_key[q]

# === Process All OMRs ===
all_results = [answer_key_row]
for filename in os.listdir(input_folder):
    if filename.lower().endswith(('.jpg', '.jpeg', '.png', '.tif')):
        filepath = os.path.join(input_folder, filename)
        image = cv2.imread(filepath)
        row = {'Image': filename}
        roll_number = detect_roll_number(image)
        row['RollNumber'] = roll_number
        student_answers = extract_answers_from_image(image)
        marks = 0
        for q in range(1, 21):
            stu_ans = student_answers[q]
            row[f'Ans_{q}'] = stu_ans
            if stu_ans == answer_key[q]:
                marks += 1
        row['Obtained Marks'] = marks
        all_results.append(row)

# === Save to Excel ===
df = pd.DataFrame(all_results)

if os.path.exists(output_file):
    try:
        os.remove(output_file)
    except PermissionError:
        print(f"❌ Please close the Excel file: {output_file}")
        exit()

df.to_excel(output_file, index=False)
print(f"✅ All data saved to: {output_file}")
