@echo off

set /p status="Have you put the file you wish to format into the same directory where this file is saved? (y/n)"
if %status%==n echo Script is terminating, please copy the file the try again & timeout /t 16 & exit /b

set /p file="What file would you like to format? e.g. myATARstuff.xlsx (remember the .xlsx)"
set /p max="What is the maximum amount of subjects students can choose (default = 7)? "
set /p fileType="Which file type are you reformatting (0 for IA, 1 for ATAR)"

xcopy /f "%file%" %USERPROFILE%\Documents

if %fileType%==0 (python %USERPROFILE%\Documents\IA_Reformat.py -fname "%file%" -smax %max%)
if %fileType%==1 (python %USERPROFILE%\Documents\ATAR_Reformat.py -fname "%file%" -smax %max%)

set destdir = %~dp0
move %USERPROFLE%\Documents\formattedATARPredictions.csv %destdir%

echo The file has now appeared in the folder. This program will automatically terminate.

timeout /t 10
exit /b
