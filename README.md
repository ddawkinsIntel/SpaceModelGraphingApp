# SpaceModelGraphingApp

Purpose:
- Graphs data from the space model

What:
- Python desktop app in the form of an exe file

How: 
- Download and save the .exe file in a folder of your choice
- Then select the .exe file and a window with the title "Space Model Graphing App" should pop up

![image](https://user-images.githubusercontent.com/89600331/151611080-8144a6eb-2d65-4c72-b9cb-ca1779a1d937.png)
- Enter a valid "Version ID" and "WIF ID" from the "Space Model" application
  -  If "WIF ID" is left blank it will default to 0
- Click the "Submit" or press "enter" to run the script
- If valid data is entered, the script will output a graph and save it in a folder named "graphs" at the same place the user saved the .exe file

package: pyinstaller main.py -F -n SpaceModelGraphApp --icon factory.ico --clean