# **Webex Message Analytics**

![Chart](https://i.imgur.com/VZ5D734.png)[^1]
[^1]: Example of user stats chart using fake names and data.

---

**Webex Analytics is a simple script that gives you an Excel spread sheet containing info on every user in a specified room.**

### Features
- Processes everything on your computer, so you don't have to worry about 3rd party companies handleing your data
- Automatically displays data in a chart
- Returns a .csv file in addition to an excel file
- Allows for searches of any date range
- 2nd page with extra stats, which are useful for measuring activity

---

### How to use

It is easist for Windows users to download webexMessageAnalytics.exe. Upon running, instructions are displayed for the user.

Alternatively, you can use webexMessageAnalytics.py. You will need the webex teams sdk (```pip install webexteamssdk```), pandas (```pip install pandas```), and xlsxwriter (```pip install xlsxwriter```) installed. From there, just run the script and all user instructions will be displayed in the console.

---

### Example output
Example user stats spread sheet
![Example output](https://i.imgur.com/aidXSjQ.png)[^2]
[^2]: Example image of user stats spread sheet using fake names and data.

Room stats spread sheet (can be found on 2nd page of excel file)

![Example stats page](https://i.imgur.com/tps4oPF.png)[^3]
[^3]: Example image of room stats page using fake data.
