# CX web search automation  
## Table of contents
- [About the program](#about-the-program)
- [Installing the program](#installing-the-program)
- [Locate the program](#locate-the-program)
- [Using the program](#using-the-program)
- [Additional points to note](#additional-points-to-note)

<a id="about-the-program"></a>
## About the program
This program is exclusively designed for **USIU-A** students to simplify course searches on the CX portal. By automating search, data scraping, and Excel sheet generation, it streamlines the course selection process. Key features include automated searches, course data extraction, excel sheet generation, and support for main filters such as Semester and faculty members.  

<a id="installing-the-program"></a>
## Installing the program
There are 2 ways to install it:

**1. Using git clone**  
Use the following command on your local terminal and wait for completion:
>git clone https://github.com/Shaurav-Vora/CX-Web-Search-Automation.git

*Note: You need to [install git](https://git-scm.com/downloads) to use this method.* 

**2. Zip file**  
Under the green button <code style="color : green">**<> code**</code>
, choose "Download zip" from the dropdown or [click here](https://github.com/Shaurav-Vora/CX-Web-Search-Automation/archive/refs/heads/main.zip) to start the download directly.

<a id="locate-the-program"></a>
## Locate the program
The program is located in the "~\CX-Web-Search-Automation\dist\main\main.exe".  

A CX.ico image is provided to customize the icon for the program but can only be applied to shortcuts of the program. Once a shortcut is created:
1. Select the shortcut and press alt+enter or right-click and select properties.
2. Choose "change icon" and browse the CX.ico. *You may also make your own icon and use that*.

<a id="using-the-program"></a>
## Using the program
![Screenshot of program](/Images/program_shot.png)

|Number      | Feature     |    Description                      |
|:-----------:|:-----------  |  :-----------------------------------|
|1           | Login       | Use your USIU login credentials here|
|2           | Filters     | Write your courses using comma separation without spaces, choose semester and faculty from dropdown and select preferred days if needed.|
|3           | Headless mode| Use this mode if you dont want the program to open the browser. This mode will execute in the background so you can continue your other work.|
|4           | Generate schedule| Click here when you are ready and an excel sheet will generate with data using the filters used|

<a id="additional-points-to-note"></a>
## Additional points to note
1. The speed of the program will depend on the speed of CX portal.
2. The program also opens a **terminal** when launched. Do **NOT** close this as it will help find errors for reporting and is also needed to use the program.
3. If you want to move the program to another folder, move the main.exe and chromedriver together. Excel sheet is not required but will create a new one wherever the files are moved to.
4. Not all errors have been captured in the program therefore incidents such as *server timeouts* may **break** the program. In such cases, **disable headless mode** and monitor the flow until the problem is found. If you wish to report it, take **screenshots** to attach and copy paste any code in the **terminal** and email it to me at shauravvora@gmail.com with the subject "CX bug".  
