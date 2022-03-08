# GTScripts.gs

This Google Apps Script integrates into Google Sheets to provide supplementary function to managing a laboratory mouse colony. This script was developed as a side project to include some basic colony management software functions.

Once installed, in the top menu of the Genotyping Sheet document, you should see a tab on the end called "GT Scripts". This menu has the functions seen below:

![image](https://user-images.githubusercontent.com/37638547/147890501-3c0d49b3-a153-4375-b3db-0133783a0c9b.png)

##### 1. Process New Tags from Form

Attached to the genotyping sheet is a Google Form:
https://docs.google.com/forms/d/1M0EjTkX0Ss2hogh8EgfWaz4sPnEiF2PP9INIsY_miE4/viewform
This form allows you to fill in all the necessary information to add new tagged pups to the GT sheet.
It does so by populating the answers into a tab called Form Responses 1:

![image](https://user-images.githubusercontent.com/37638547/147890507-787a88ef-f68e-4cbc-bc00-1e49fd0768a2.png)

This script automatically parses this information and places it into the appropriate places of the appropriate genotyping sheet, and then removes the form result entry.

FOR THIS TO WORK CORRECTLY, you need to enter mouse ranges into the form as 4-digit numbers (i.e. 1234-1238). Single numbers are supported, as are non-contiguous ranges separated by dashes and commas (i.e. 1234,1236,1238-1240,1241-1242,4444). If none of one gender, leave blank. Try not to feed it impossible ranges; I haven't tried that.

Note: the form populates the sheet from the top, and the script reads from the bottom.
	After everything's added, alphabetize the sheet by column A to get numbers in sequential order.

Usage:
	1) Fill and submit the Google form.
	2) Press the Process button.

##### 2. Update WAB Total #s

A very simple method, this reads the total # of living mice from the coversheet, and appends it with the current date to column C of the data sheet.  This live-updates the graph on the coversheet.
	If you're adding new tails and processing sacrifices, do everything before running this.
	Run at least once a week for historical record purposes.

Usage:
	1) Change the # of mice marked as alive (add and/or sac, scripted and/or manually).
	2) Press the Update button.

Mouse Management Tools

##### 3. Move a Mouse to Thorn

This allows you to manage the status of Thorn doxycycline-induced mice easier. When you plug in a number, it searches the spreadsheet for the mouse, marks it as sacrificed and hides the row, but copies the # and full genotype string to the Thorn sheet, along with today's date.

![image](https://user-images.githubusercontent.com/37638547/147890521-7954a08f-f1d3-4eef-ab18-2e0536a513b5.png)

This sheet has room for comments, including doxy and death dates.
Feel free to hide rows when you're done.

You can enter a single number, or a list of numbers separated by commas.

Usage:
	1) Set up mice to transfer to Thorn, submit transfer paperwork.
	2) Press the Move button, enter mouse # or #s (separated by commas).
	3) Manage details like doxy & death dates on the Thorn tab.

##### 4. Sacrifice a Mouse

This method removes the Alive flag, adds a sac comment to the end of the comment field, and hides the row of the specified mouse.

You can enter a single number, or a list of numbers separated by commas.

Usage:
	1) Keep track of every mouse # you sac in WAB.
	2) Press the Sacrifice button, enter mouse # or #s (separated by commas).

##### 5. Hide Dead Mice

This method should iterate through the active sheet and hide any rows missing the Alive flag.
This sometimes doesn't work. The code is overall correct, but the system can crash midway through iterating.
	The issue may be the result of the iteration variable exceeding the max row #.
	"Service error: Spreadsheets"
I've recently had some success with it, just waiting for a while.

Usage:
	1) Push the button, WAIT.  If it crashes, hide manually.

##### 6. Get Mouse Info

This method provides an info card on the inputted mouse #.
It includes the #, an A if alive, the sex, birthdate, genotype, parentage (MOMxDAD), and any notes.

![image](https://user-images.githubusercontent.com/37638547/147890527-f4d30149-2bc9-499d-aa55-9855c5f3bd42.png)

For older entries that use the Excel date format, it will give you a 1969 date.
	Converting the birthday column should fix that.
