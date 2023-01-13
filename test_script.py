
# first install
# pip install pypiwin32 or pip install comtypes

try:
    import win32com.client as w 
except ImportError:
    from pip._internal import main as pip 
    pip(['install', '--user', 'pywin32']) 
    import win32com.client as w 

import time as t

w.Dispatch('Illustrator.Application')
myApp = w.GetActiveObject('Illustrator.Application')

if myApp.Documents.Count==0:
    myDoc = myApp.Documents.Add(1,1920,1024) 
else:
    myDoc = myApp.Documents(1)

myDoc.Layers[0].Name = 'test'
t.sleep(0.5)
myDoc.PathItems.Rectangle(160,100,200,100) 
t.sleep(0.5)
myDoc.PathItems.RoundedRectangle(850,100,200,100, 50)
t.sleep(0.5)
myDoc.PathItems.Ellipse(200,150,100,100)
t.sleep(0.5)
# myDoc.PathItems.Polygon(200,500,75,2, 8)
# t.sleep(0.5)
myDoc.PathItems.Star(200,300,10,75,6)

Col = w.Dispatch("Illustrator.RGBColor")
Col.red = 0
Col.Blue = 100 
Col.Green = 100

Col2 = w.Dispatch("Illustrator.RGBColor")
Col2.red = 0
Col2.Blue = 200 
Col2.Green = 200
t.sleep(0.5)

for i in range(myDoc.PathItems.Count):
     myDoc.PathItems[i].fillColor = Col2
     myDoc.PathItems[i].strokeWidth = 3
     myDoc.PathItems[i].strokeColor = Col
     t.sleep(0.5)


# python test_script.py



''' 

Illustrator Scripting in Python


Scripting in Illustrator is used to automate repetitive tasks and are often used as a creative tool to streamline tasks that might be too time consuming to do manually. For example, you could write a script to generate a number of localized versions of a particular image or to gather information about the various color profiles used by a collection of images.

Illustrator COM & DOM
Illustrator can be scripted through COM(Component Object Model). Its DOM(Document Object Model) is the same when accessing it through either its own JavaScript engine or Python or any other scripting language it supports. The Illustrator DOM consists of a hierarchical representation of the Illustrator application, the documents used in it, and the components of the documents. The DOM allows you to programmatically access and manipulate the Artboard and its components. For example, through the DOM, you can create a new document, add a layer to an existing document, or change the background color of a layer. Most of the functionality available through the Illustrator user interface is available through the DOM.

But why Python?
Illustrator scripting officially supports JavaScript, AppleScript & VBScript. However, scripting in Python is also fairly easy if not easier if you're already comfortable with Python. You may have already heard that Python is gaining in popularity, but did you know it’s now the most popular introductory programming language in U.S. universities? Python is also cross platform just like JavaScript is and lately becoming one of the fastest growing programming language according to StackOverflow as of 2017 / as of 2019

Python is easy to use, powerful, and versatile, making it a great choice for beginners and experts alike. Python’s readability makes it a great first programming language - it allows you to think like a programmer and not waste time understanding the mysterious syntax that other programming languages can require.

Getting Started
Python allows you to access COM and it's DOM with the help of a Python extensions like "pypiwin32" or "comtypes". Install these modules and you're ready to start scripting Illustrator in Python

pip install pypiwin32 or pip install comtypes
Hello World!
from win32com.client import GetActiveObject

app = GetActiveObject("Illustrator.Application")
docRef = app.Documents.Add()
rectRef = docRef.PathItems.Rectangle(700, 50, 100, 100)
areaTextRef = docRef.TextFrames.AreaText(rectRef)
areaTextRef.Contents = "Hello World!"
How to inspect scripting object properties?
There's not a straight forward way, you need to read the documentation to understand what properties/attributes are available for a scripting object, or possibly a COM browser. For example, I've extracted the Python scripting object reference for Illustrator CC 2018 at api_reference

Scripting on Mac?
Yes, scripting on Mac is also possible, see photoshop_mac_scripting for more details as a reference to getting started

Illustrator Scripting Resources
Illustrator Scripting Resources
Illustrator Scripting Guide
Illustrator Scripting Developer Forum
Illustrator Scripting Javascript API Reference
Illustrator Scripting Javascript Tutorials
Also see
InDesign Scripting in Python
Photoshop Scripting in Python
Contribution
If you've written a useful Illustrator Python script and wants to share with the world, please create a new issue with the file as an attachment to the issue.

When you submit a script, please try to include the following information at the start of your script

# script_file_name.py

# Created: 1st January 2019
__author__ = 'Your Name or Original Author Name'
__version__ = '1.0'

"""
A short description of what the script does
"""

"""
Instructions on how to use the script, if any
"""
Go to illustrator-scripting-python/issues/new
Add title as Useful Script
Drag & drop your .py script file into the description area
Click Submit new issue

 '''