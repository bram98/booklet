
import glob
import os
import shutil
from wand.image import Image as wImg
from wand.display import display

src_folder = './response_data/portraits_renamed/'
portraits = glob.glob(src_folder + '*')