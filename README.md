# SpaceModelGraphingApp

Purpose:
- Graphs data from the space model

What:
- Python desktop app in the form of an exe file

WARNING:
- MUST save "ConfigFile.xlsx" to the same location as the .exe file after downloading the .exe file
- MUST have an excel file open when running the app
- Make sure you have already solved for the scenario you are trying to graph


<h2>How to:</h2>

1. Download and save a copy of the .exe file, located <a href="https://intel.sharepoint.com/:f:/r/sites/tmgspcapdesign/Shared%20Documents/Space?csf=1&web=1&e=wVSzW6">here</a>, in a folder of your choice (preferably to a location that is easy to find)
2. Save a copy of the "ConfigFile.xlsx", located <a href="https://intel.sharepoint.com/:x:/r/sites/tmgspcapdesign/Shared%20Documents/General/Documentation/SpaceModelGraphingAppFiles/ConfigFile.xlsx?d=w478c37764b624fc78146d3dcf512b5e3&csf=1&web=1&e=bMZaQs">here</a>, in the same folder as the .exe file
3. Populate "ConfigFile.xlsx" with your UID (ex. AMR\ddawkins)
    - DO NOT alter the Server & Database values, unless Uma or ZZ say the Server or Database values have changed
    - Save and close this file after making changes
4. Then double clike the .exe file and a window with a black screen will appear then a window with the title "Space Model Graphing App" will pop up

![image](https://user-images.githubusercontent.com/89600331/151611080-8144a6eb-2d65-4c72-b9cb-ca1779a1d937.png)

5. Enter a valid "Version ID" and "WIF ID" from the "Space Model" application
    -  If "WIF ID" is left blank it will default to 0
6. Click the "Submit" or press "enter" to run the script
7. If valid data is entered, the script will output a graph and save it in a folder named "graphs" at the same place the user saved the .exe file
8. (Optional) If you'd like to customize the colors for the processes use the SpaceModelColorMap.xlsx template, which can be downloaded <a href="https://intel.sharepoint.com/:x:/r/sites/tmgspcapdesign/Shared%20Documents/General/Documentation/SpaceModelGraphingAppFiles/SpaceModelColorMap.xlsx?d=w7bb6d87991c54a239b1618a2035be819&csf=1&web=1&e=zF2LYW">here</a>, and save it in the same location as the .exe file
    - If you use the choose to use the template, you MUST declare a color for all process that are being used, else a default color will be used
    - Save and close this file after making changes

<h2>Troubleshooting:</h2>

- Getting an Error? Make sure you have an excel file open, the file can be blank, when running the app
- Permission Denied? Make sure to save and close the "ConfigFile" and "SpaceModelColorMap" file after making changes
- All graphs are blank? Make sure you solved for the scenario you are trying to graph
- If you add or remove a process after initially making the graphs, then you MUST delete the original graphing file PRIOR to re-running the graphs
- For any other issues, 
    1. When in doubt delete the original output file
    2. Solve the space model scenario 
    3. Re-run the graphing app


package: pyinstaller main.py -F -n SpaceModelGraphApp --icon factory.ico --clean
