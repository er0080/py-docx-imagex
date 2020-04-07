#!/usr/bin/python3

import os
import sys
import argparse
from pathlib import Path
import zipfile as zf
import uuid
from wand.image import Image
import subprocess

def main(argv):
    
    parser = argparse.ArgumentParser(description='Extract images from MS Word .docx file(s) and save in JPEG format')
    
    parser.add_argument("infiles", metavar = "file", type = argparse.FileType(), 
                        nargs="+", help="input .docx file(s)")
    
    parser.add_argument("-o", "--outdir", type = Path, 
                        #default = Path(__file__).absolute().parent, 
                        default = Path.cwd(),                        
                        help="output directory for image files")

    args = parser.parse_args()

    # check for directory existence and create/error out as needed
    if args.outdir.exists() and not args.outdir.is_dir():

        print ("output directory name already exists as a file, please rename")
        exit(1)

    elif not args.outdir.exists():
        args.outdir.mkdir()

    for infile in args.infiles:

        try:
            indocx = zf.ZipFile(infile.name)
        
        except zf.BadZipfile:
            print (infile.name, "is an invalid / corrupt .docx file")
            exit(1)
        
        for f in indocx.namelist():
            if "word/media/" in f:
                with indocx.open(f) as mediafile:
                    print("processing", mediafile.name, "from", infile.name, "...")
                    
                    # Rename file w/format: x-image-uuid.jpeg as x in x.docx:
                    outfile = args.outdir / str(infile.name + "-image-" + str(uuid.uuid4()) + ".jpeg")

                    # test if jpeg format and one-to-one copy:
                    if (".jpeg" or ".jpg") in f:                
                        print ("writing to", outfile)
                        
                        # pathlib.Path.write_bytes takes care of opening & closing file...                
                        outfile.write_bytes(mediafile.read())                    
                    
                    elif ".png" in f:
                        print ("converting .png to", outfile)
                        
                        im = Image(file=mediafile)
                        im.format = 'jpeg'
                        outfile.write_bytes(im.make_blob()) 
                    
                    elif ".emf" in f:
                        print ("converting .emf to", outfile)
                    
                        # need to use unoconv, then ImageMagick, also fu@$!in tempfiles again!
                        tmp = args.outdir / "tmp_file.emf"
                        tmp.write_bytes(mediafile.read())

                        pdf = subprocess.run(["unoconv", "--format=pdf", "--stdout", tmp.resolve()],
                                              capture_output=True, check=True)
                        
                        jpeg = subprocess.run(["convert", "-density", "150", "-trim",
                                               "-bordercolor", "white", "-border", "5", "-", "jpeg:-"], 
                                               capture_output=True, input=pdf.stdout,
                                               check=True)                    

                        outfile.write_bytes(jpeg.stdout)
                        tmp.unlink()

                    elif ".wmf" in f:
                        print ("converting .wmf to", outfile)
                    
                        # need to use unoconv, then ImageMagick, also fu@$!in tempfiles again!
                        tmp = args.outdir / "tmp_file.wmf"
                        tmp.write_bytes(mediafile.read())                       
                        
                        pdf = subprocess.run(["unoconv", "--format=pdf", "--stdout", tmp.resolve()],
                                              capture_output=True, check=True)
                        
                        jpeg = subprocess.run(["convert", "-density", "150", "-trim",
                                               "-bordercolor", "white", "-border", "5", "-", "jpeg:-"], 
                                               capture_output=True, input=pdf.stdout,
                                               check=True)                    

                        outfile.write_bytes(jpeg.stdout)
                        tmp.unlink()              
                                       
                    elif ".wdp" in f:
                        print ("converting .wdp to", outfile)
     
                        #ImageMagick JXR handler needs an actual file; won't take bytestream...
                        tmp = args.outdir / "tmp_file.wdp"
                        tmp.write_bytes(mediafile.read())
                     
                        im = Image(filename=tmp.resolve())
                        im.format = 'jpeg'

                        #im.save(filename=outfile.name)                    
                        outfile.write_bytes(im.make_blob())

                        #delete tmp_file.wdp file
                        tmp.unlink()
                
                    else:
                        print(mediafile.name, "is not a recognized image type")

                    mediafile.close()
                    

if __name__ == "__main__":
    main(sys.argv[1:])

