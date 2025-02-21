# Setup

## Installing Python
**NOTE**: If you already have Python installed, ensure that it is above version 3.7. This version was selected because some packages may not work as expected with a lower version. If it is above version 3.7, you may skip this step.<br />

If you do not have Python installed, you should use the provided installer (or don't, up to you, you may also install it from the Python website).

This will ensure that you have a sufficient version, and it is the one I personally use when coding.
Ensure that you check the box that says **Add Python to PATH** at the bottom of the installer, this is extremely important and can cause issues with the setup file if not done right.

## Running the setup file
There is a setup file provided, which will automatically install all the required packages. To run this, simply double click the file called "setup" in the folder.

If everything goes smoothly, a run file will be created, which you can double click to run the program with ease.

# Running the program
Running the program is very simple. But here's a walkthrough on how to do it.

**NOTE**: Each step of this guide assumes you are pressing Enter when prompted.

## Starting the program
The first step before the program is to actually run it. For ease of use, you may simply double click the **run** file in the same folder as everything.

## Selecting a file
The first step of the program will be to select the spreadsheet file for the well data. It will open up a file dialog, from there you can select the spreadsheet you want to use.

## Selecting a range
If everything goes well, you'll be given a list of dates to choose from, assuming the data is not malformed.

You will be asked for four inputs, the start date, start time, end date, and end time.

Keep in mind that the times are inclusive, meaning that if something happens at 12:30, and you start at 12:30, it will be included.
- When asked for a date, input the number **next** to the date you want to select.
- When asked for a time, follow a 24-hour format of HH:MM. You may simply leave any time fields blank to automatically start/end at midnight.

## Understanding output
From here, it is very straightforward. If there were any issues, it would be logged to the terminal.


Keep in mind that warnings are not always bad, if there is data missing where it usually goes, it will send a warning, but this can be expected.

Always be sure to double check your datasets if there is a warning, but most of the time they can be ignored.

## Saving the report
Another file dialog will be opened, but this time it will ask you to save a file. This will overwrite the file, meaning it will replace all the data in the selected file.

**This is NOT the full report.**  This will only create a document with a similarly formatted table for the time range provided. The formatting is not exactly 1:1, so you might not be able to copy paste it, but it should have the correct data.

## End menu
After the end of the program, you will be given two options. You may exit the program, or generate another report, with the option of using the same file or picking another one.

This is pretty simple, just type the number next to the option and press Enter. Generating another report will simply take you straight back to the beginning.

Thank you for using my program!

# Troubleshooting
## During setup
### Python is not installed

There are three main reasons this could happen.
1. **Python really isn't installed**
   
	To fix this, run the Python installed provided in the folder, or go to https://www.python.org/downloads/ to install a version of Python newer than version 3.7.

2. **Incorrect installation settings**

	If Python is installed, but it is not added to PATH. To fix this, you don't have to reinstall Python, simply run the installer and click **Modify**. After this, you may click **Next** at the bottom, then check the box that says **Add Python to environment variables**.

3. **Restart required**

	After installing Python, a restart may be required to run Python files through the setup script. Simply restart your computer and try again.

### Python version must be at least 3.7
This is because you may already have Python installed, and it is too old for the program to use. To fix this, run the given installer or install a version from the Python website that is newer than 3.7. This may require you to uninstall your current version, which you can do by searching up Python in your computer's applications, and uninstalling from there.

### Failed to install requirements
This one may be a bit problematic, since this issue can cause significant problems and requires additional help.
If you know what you are doing, you may try to type this into your terminal in the directory of the program: `pip install -r requirements.txt`

If this does not fix it, or yields further errors, please see the end of the page.

## During runtime
This section is still a work in progress!

### There's a bunch of weird symbols everywhere
This may be due to your terminal being relatively primitive, and not accepting color codes. This might be annoying, but it's visual only!

### There's a bunch of random words on my screen
These are most likely debug messages from the file dialogs. These cannot be disabled, but you have nothing to worry about.

### Something pops up and then disappears
If the script doesn't run, and something pops up for a second before disappearing, this is because an error is occurring and is not being caught. To resolve this, see the end of the page.

## Still need help?
This section is still a work in progress!
