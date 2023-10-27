vs_scanwatch

The function of this service is to scan the contents of a specified folder and attempt to move it's contents to a subfolder inside this folder. 
If the subfolder does not exist, the service will create it.
This function is to be executed regularly after a specified amount of time.

App.config keys explained

"overrideFolderPath" - this key controls the root folder path    
The default value is C:\Scan\.
If a valid path to an existing folder is provided the key will override the default value. 
If a valid path is provided but the folder does not exist, the default root folder path will be used. 
In the case that the default root folder path is used and the C:\Scan\ folder does not exist,
the service will create a C:\Scan\ folder to ensure operation.
Values of "false", "False" or "FALSE" will instruct the app to use the default value.
 
"overrideDate" - this key controls the output folder behaviour
The default behaviour is to set the folder to the current date in Belgrade in yyyyMMdd format
If a valid date in the format yyyyMMdd is provided, the key date will be used instead

"scanInterval" - this key controls the time interval that passes before the service attempts to execute it's main programming again
The default value is 30000. The value represents miliseconds, so 30000 equals to a wait time of 30 seconds.
The value entered needs to be a valid number or else the default value will be used.
The minimal wait time is 10 seconds(10000), if a lower value is entered, the default value will be used.
Values of "false", "False" or "FALSE" will instruct the app to use the default value.

"debug" - this key controls wether debug logging is enabled or disabled
Values of "true", "True" or "TRUE" will enable debug logging inside a separate debug log file instead of the regular file.
    