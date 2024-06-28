# EVED-SystemMonitoringAutomation

## Setup 

1. DO NOT 
    - Please don't autosave/save to one drive. Please ensure your file is saved locally on your Laptop/PC
    - Please don't Include spaces in the excel file (Renaming is fine)
    - Please don't Rearrange sheets in file (Renaming is fine, if you like to add sheets please add after "Charts3" or Sheet 4)

2. In excel follow the instructions below
    - 1 - File > Options > Trust Center > Trust Center Setting > Macro Settings > Select the following: "Enable VBA Macros", "Enable Excel 4.0 macros", "Trust access to the VBA project object model"

    - 2 - (Now back in the excel spreadsheet) Right click "Home" ribbon > Customize the ribbon > Import/Export > Import customization file > Select "EVEDsystemUI"

    - 3 - (use Alt + Fn + F11 or Fn + F11 to open VBA editor) On This Screen Go To The Top Menu Bar > Tools > References. Then Ensure You Have Your References Checkboxed And Matching The References In The Image Below. (All Unchecked References Are In Alphabetical Order To Help Find Them)

    ![alt text](https://github.com/Ross-Ede/HomeownerClassAutomation/blob/main/images/image-4.png?raw=true)

    - 4 - Use Ctrl + S to save your changes so far (You can now close out the VBA editor at this point)

## Instructions on use
    1. In "DataAnalyzed" or Sheet 1 append this months recent data onto the bottom (adjacent to last months data, please do not leave any space)
        * Note that the program uses the system date to match monthly updates. So long as your systems month and year match the data's month and year the program will update.

    2. If the custom ribbon was imported properly you should see a "UPDATE" tab/ribbon atop in your excel now. From there you can select "Run All" which will update the entire files charts, averages, data, formatting etc.
