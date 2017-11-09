import os
import glob

files = glob.glob('r/IMG*.jpg')

for file in files:
  os.rename(file, 'r/KAPTA_IMAGENES_{}'.format(file.split('_')[1]))
