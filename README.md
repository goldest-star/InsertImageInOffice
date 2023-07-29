InsertImage
======

InsertImage is a Microsoft Office add-in which makes Word, Excel, and PowerPoint support most image file formats (via ImageMagick convert).

- Support all MS Office products: Word, Excel, and PowerPoint.
- Support MS Office 2003 to Office 2016.
- Mac OS Support is also in massive development.

# Prerequisite
You may need to install these packages first:
- [Ghostscript](https://www.ghostscript.com/download/gsdnld.html) for PDF, PS file processing

# Quick Start

### Install
1. `git clone https://github.com/goldest-star/InsertImageInOffice.git`
2. `cd InsertImagePlus/`
3. `.\setup.bat`

### Usage

Select "Insert/Change Image" from the "Add-Ins" tab of the ribbon, and you will get a file dialog where you can load your image:

![alt text](https://github.com/goldest-star/InsertImageInOffice/blob/main/Office_Add-Ins.png)

Choose your image, and click on "Open". InsertImage++ will convert your image into PNG format, and insert it into Office.

If you need to change the image, just select the image, then click on "Insert/Change Image" again, and the file dialog will re-appear so you can load another image.

You can also treat the loaded image as an ordinary PowerPoint image. For example, it can be grouped, animated, rotated, moved, and resized. Further editing of the loaded image will preserve all these changes.
