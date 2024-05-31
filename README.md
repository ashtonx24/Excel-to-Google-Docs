# Excel-to-Google-Docs
With this code, you can use an excel sheet to input data onto a google form.

This readme is for windows users but you may use chatgpt for usage intrustions for mac or linux ;)
You need to have the following things on your computer or laptop for this to work:
- python idle
- google driver for your chrome version
- a few python libraries installed using the pip command which will be declared later on.

First you have to view your chrome version which can be done by going to the 3 dots on the top right corner, then help (can be found at the bottom of the list), then about chrome.

Use that and perform a google search for a google driver for your chrome version.

Once you have found and installed it, you will haev a zip file downloaded. Unzip it into a folder in your main directory and copy its path.

Then, you will have to go to your system's environment variables and add that path there. Once done, you are good to go.

Now comes the libraries. Just run the following commands on the console of your idle or your command prompt and you are good to run the code.
'pip install pandas openpyxl selenium'
Don't forget to remove the inverted commas!

Now, for the Xpaths that are in the code. Just open your google form, right click on the entry fields and click on inspect. A part will light up/ get highlighted on the screen that pops up, just right click on that and click on copy as XPATH.
Then, you can replace easily.
