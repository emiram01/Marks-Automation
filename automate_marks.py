#!/usr/bin/env python
from openpyxl.drawing.image import Image
import argparse
import datetime
import subprocess
import openpyxl
import pymongo
import sys
import csv
import os
import re

# Handle arguments
parser = argparse.ArgumentParser()
parser.add_argument('--files', '-f', dest='work_files', nargs='+', type=str, help='work files to process')
parser.add_argument('--xytech', '-x', dest='xytech_file', type=str, help='xytech file to process')
parser.add_argument('--output', '-o', dest='output_type', type=str, help='output to DB or CSV or XLS')
parser.add_argument('--process', '-p', dest='video_file', type=str, help='video file to process')
parser.add_argument('--verbose', '-v', action='store_true', help='show verbose')
args = parser.parse_args()

if not any(vars(args).values()):
    print("""Please provide an argument:
    --files: work files to process
    --xytech: xytech file to process
    --output: output to DB or CSV or XLS
    --process: video file to process
    --verbose: show verbose"""
    )
    sys.exit(2)

# Handle database
client = pymongo.MongoClient("mongodb://localhost:27017/")
db = client["marksAutomationDB"]
requestLogCol = db["requestLogs"]
workFileCol = db["workFiles"]
if args.verbose: print("DATABASE INFO:", client, db, requestLogCol, workFileCol)

# Process video
if args.video_file is None:
    print ("No video file selected")
    sys.exit(2)
else:
    video_file_location = args.video_file

def get_framerate():
    cmd = ['ffprobe', '-v', 'error', '-select_streams', 'v', '-show_entries', 'stream=r_frame_rate',
           '-of', 'default=noprint_wrappers=1:nokey=1', video_file_location]
    result = subprocess.run(cmd, capture_output=True, text=True)
    framerate = int(eval(result.stdout))
    return framerate

def get_video_length():
    cmd = ['ffprobe', '-v', 'error', '-show_entries', 'format=duration', '-of',
           'default=noprint_wrappers=1:nokey=1', video_file_location]
    result = subprocess.run(cmd, capture_output=True, text=True)
    duration = float(result.stdout)
    return duration

def get_timecode(frame, fps):
    hours = frame // (3600 * fps)
    minutes = (frame // (60 * fps)) % 60
    seconds = (frame // fps) % 60
    frames = frame % fps
    timecode =  f"{hours:02d}:{minutes:02d}:{seconds:02d}:{frames:02d}"
    return timecode

# Video file properties
fps = get_framerate(video_file_location)
video_length = get_video_length(video_file_location)
video_frame_length = int(video_length * fps)

# Fetch marks that fall into the length of video
less_than_vid_length = []
docs = workFileCol.find({})
for doc in docs:
    loc_frames = doc["location/frames"]
    for lf in loc_frames:
        nums = re.findall(r'\d+', lf)
        frames = int(nums[-1]) if nums else None
        if frames <= video_frame_length:
            less_than_vid_length.append(lf)

locations, ranges = [], []
for string in less_than_vid_length:
    s = string.split(' ')
    if('-' in s[1]):
        locations.append(s[0])
        ranges.append(s[1])

# Convert frame ranges to timecode format
timecodes = []
for rng in ranges:
    r = rng.split('-')
    start = get_timecode(int(r[0]), fps)
    end = get_timecode(int(r[1]), fps)
    timecodes.append(start + ' - ' + end)

# Generate thumbnails for frame ranges
thumbnail_folder = './outputfiles/thumbnails'
thumbnails = [file for file in os.listdir(thumbnail_folder)]
if len(thumbnails) < len(ranges):
    if args.verbose: print("GENERATING THUMBNAILS")
    thumbnail_width = 96
    thumbnail_height = 74
    for rng in ranges:
        r = rng.split('-')
        start = int(r[0])
        end = int(r[1])
        desired_frame = (start + end) // 2
        desired_seconds = desired_frame / fps

        cmd = [
            'ffmpeg',
            '-i', video_file_location,
            '-vf', f'scale={thumbnail_width}:{thumbnail_height}',
            '-ss', str(desired_seconds),
            '-vframes', '1',
            os.path.join(thumbnail_folder, f'thumbnail_{rng}.jpg')
        ]

        subprocess.run(cmd)
        
# Match thumbnail with correct frame range
def sorting_key(tn):
    start, end = tn.split('_')[-1].split('.')[0].split('-')
    start = int(start)
    end = int(end)

    for i, r in enumerate(ranges):
        r_start, r_end = r.split('-')
        r_start = int(r_start)
        r_end = int(r_end)

        if start == r_start and end == r_end:
            return i
        
    return len(ranges)

thumbnails = sorted(thumbnails, key=sorting_key)

# Open Xytech file
if args.xytech_file is None:
    print ("No Xytech file selected")
    sys.exit(2)
else:
    xytech_file_location = args.xytech_file
    
xytech_folders = []
read_xytech_file = open(xytech_file_location, "r")
prev_line = ""
for line in read_xytech_file:
    if 'Producer:' in line:
        producer = line.split(':')[1].strip()
    if 'Operator:' in line:
        operator = line.split(':')[1].strip()
    if 'Job:' in line:
        job = line.split(':')[1].strip()
    if 'Notes:' in prev_line:
        notes = line.strip()
    if "/" in line:
        xytech_folders.append(line)
    prev_line = line

# Open work files
if args.work_files is None:
    print ("No BL/Flame files selected")
    sys.exit(2)
else:
    work_file_locations = args.work_files

output = []
for work_file_location in work_file_locations:
    curr_type = work_file_location.split('_')[0].strip()
    output.append (work_file_location)
    if args.verbose: print(work_file_location)
    read_work_file = open(work_file_location, "r")

    # Read each line from work file
    for line in read_work_file:
        line_parse = line.split(" ")
        if curr_type == 'Baselight':
            current_folder = line_parse.pop(0)
            sub_folder = current_folder.replace("/images1/Avatar", "")
        elif curr_type == 'Flame':
            current_folder = line_parse.pop(1)
            sub_folder = current_folder.replace("/Avatar", "")
        
        new_location = ""
        # Folder replace check
        for xytech_line in xytech_folders:
            if sub_folder in xytech_line:
                new_location = xytech_line.strip()
        first=""
        pointer=""
        last=""
        for numeral in line_parse:
            # Skip <err> and <null>
            if not numeral.strip().isnumeric():
                continue
            # Assign first number
            if first == "":
                first = int(numeral)
                pointer = first
                continue
            # Keeping to range if succession
            if int(numeral) == (pointer+1):
                pointer = int(numeral)
                continue
            else:
                # Range ends or no sucession, output
                last = pointer
                if first == last:
                    output.append ("%s %s" % (new_location, first))
                    if args.verbose: print ("%s %s" % (new_location, first))
                else:
                    output.append ("%s %s-%s" % (new_location, first, last))
                    if args.verbose: print ("%s %s-%s" % (new_location, first, last))
                first= int(numeral)
                pointer=first
                last=""
        # Working with last number each line 
        last = pointer
        if first != "":
            if first == last:
                output.append ("%s %s" % (new_location, first))
                if args.verbose: print ("%s %s" % (new_location, first))
            else:
                output.append ("%s %s-%s" % (new_location, first, last))
                if args.verbose: ("%s %s-%s" % (new_location, first, last))

# Handle output to either DB, CSV, or XLS
if args.output_type is None:
    print ("No output type selected")
    sys.exit(2)

elif args.output_type == 'DB':
    requestLogs, workFiles = [], []
    for i, work_file_location in enumerate(work_file_locations):
        user = os.getlogin()
        machine = work_file_location.split('_')[0].strip()
        file_user = work_file_location.split('_')[1].strip()
        file_date = work_file_location[:work_file_location.index('.')].split('_')[2].strip()
        submitted_date = datetime.date.today().strftime("%Y%m%d")

        if i < len(work_file_locations) - 1:
            frames = output[output.index(work_file_location) + 1: output.index(work_file_locations[i+1])]
        else:
            frames = output[output.index(work_file_location) + 1: ]

        requestLogs.append({"user": user, "machine": machine, "file_user": file_user, "file_date": file_date, "submitted_date": submitted_date})
        workFiles.append({"file_user": file_user, "file_date": file_date, "location/frames": frames})

    requestLogCol.insert_many(requestLogs)
    if args.verbose: print ("REQUEST LOGS:", requestLogs)
    
    workFileCol.insert_many(workFiles)
    if args.verbose: print ("WORK FILES:", workFiles)

elif args.output_type == 'CSV':
    with open('./outputfiles/output.csv', 'w',  newline="") as csvf:
        csvw = csv.writer(csvf, delimiter=',')
        fields = ['Producer', 'Operator', 'Job', 'Notes']
        field_values = [producer, operator, job, notes]
        csvw.writerow(fields)
        csvw.writerow(field_values)
        csvw.writerow([' '])

        for line in output:
            csvw.writerow([line])

elif args.output_type == 'XLS':
    workbook = openpyxl.Workbook()
    sheet = workbook.active
    fields = ['Location', 'Frames', 'Timecodes', 'Thumbnails']
    sheet.column_dimensions['A'].width = 60
    sheet.column_dimensions['B'].width = 12
    sheet.column_dimensions['C'].width = 25
    sheet.column_dimensions['D'].width = 20
    sheet.append(fields)

    for i in range(len(ranges)):
        sheet.row_dimensions[i + 2].height = 75
        image = Image(thumbnails[i])
        sheet.add_image(image, f'D{i + 2}')
        sheet.append([locations[i], ranges[i], timecodes[i]])
    
    workbook.save('./outputfiles/output.xls')
    