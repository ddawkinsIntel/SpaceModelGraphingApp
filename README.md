# SpaceModelGraphingApp

Purpose:
- Graphs data from the space model

What:
- Python desktop app in the form of an exe file

WARNING:
- MUST have an excel file (it can blank) open when running the app
- Make sure you have solved for the scenario you are trying to graph, or 

How: 
- Download and save the .exe file in a folder of your choice
- Then select the .exe file and a window with the title "Space Model Graphing App" should pop up

![image](https://user-images.githubusercontent.com/89600331/151611080-8144a6eb-2d65-4c72-b9cb-ca1779a1d937.png)
- Enter a valid "Version ID" and "WIF ID" from the "Space Model" application
  -  If "WIF ID" is left blank it will default to 0
- Click the "Submit" or press "enter" to run the script
- If valid data is entered, the script will output a graph and save it in a folder named "graphs" at the same place the user saved the .exe file

Troubleshooting:
- MUST have an excel file (it can blank) open when running the app
- All graphs are blank, make sure you solved for the scenario you are trying to graph
- If you add or remove a process after initially making the graphs, then you MUST delete the original graphing file PRIOR to re-running the graphs
- For any other issues, 
    1. When in doubt delete the original file
    2. Solve the space model scenario 
    3. Re-run the graphing app


package: pyinstaller main.py -F -n SpaceModelGraphApp --icon factory.ico --clean
