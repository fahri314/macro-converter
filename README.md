https://macroconverter.com

![](https://github.com/fahri314/macro-converter/raw/master/static/images/logo1.png)

~ Macro conversion project from Excel macro to LibreOffice macro ~

# Overview

The purpose of the existence of this project, Microsoft Excel workbooks macro codes doesn't works on the libreoffice calc.

This is because, Excel Visual Basic for Application (VBA) and LibreOfice Basic 
 syntactically same but there are differences between VBA and LibreOfice Basic 
 object model.

A complete conversion is not possible because a method has hundreds of properties.
This project can only achieve a certain level of success.

We are waiting for your help. If you find a code match, send it to us. We will review matching and add at the project.

![](https://github.com/fahri314/macro-converter/raw/master/static/video/macro-converter.gif)

# Benefits

The user is assumed to know the VBA language and is familiar with the Microsoft Excel Object Model. The users will while trying any VBA macros learn to equivalent code on LibreOffice Basic.

Microsoft Exel VBA always updating and have very large structure. Our purpose is try to be more efficient by translating the most commonly used methods.

The user should feel free to contact the author to suggest areas to expand this document.

# Progress

![](https://github.com/fahri314/macro-converter/raw/master/static/progress/progress1.PNG)

# Used Technologies

This project was created with open source software.

[Django](https://www.djangoproject.com/)

Django is a high-level Python Web framework that encourages rapid development and clean, pragmatic design. Built by experienced developers, it takes care of much of the hassle of Web development, so you can focus on writing your app without needing to reinvent the wheel. It’s free and open source.

[Python](https://www.python.org/)

Python is developed under an OSI-approved open source license, making it freely usable and distributable, even for commercial use. Python's license is administered by the Python Software Foundation.


# Requirements

    Pyhon(3.6) (Recommended)
    Django(2.0.5) (Recommended)
# Smtplib Settings

Change mail settings:

    -/macro    
      -/send_mail.py
        username = "your@mail.com"
        password = "yourpassword"
        receiver = "receiver@mail.com"

# Security

If you want upload any django project to server, close debug mode on settings.py
    
    -/macro_converter
      -/settings.py
        DEBUG = False
        ALLOWED_HOSTS = ['127.0.0.1']   # Your domain or ip adress

# License

This project is licensed under the terms of the <b>GNU General Public License v3.0.</b>

![](https://github.com/fahri314/macro-converter/raw/master/static/images/gplv3.png)

For more details, check this link: [gnu.org](https://www.gnu.org/licenses/gpl-3.0.html)

[Licence.txt](https://www.gnu.org/licenses/gpl.txt)
