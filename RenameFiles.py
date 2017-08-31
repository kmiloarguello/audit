import exifread

f = open('IMG_1174.JPG', 'rb')
tags = exifread.process_file(f)

print tags