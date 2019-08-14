@echo OFF
:: @author Robert Todar <robert@robert.todar.com>
:: This is a simple template for a basic CLI using just a batch file.
:: To use in commmand window, set your `Path` Environment Variable to include the folder path where this file exists.

:: Capture script path & filename for ease of use.
set path=%~dp0
set fileName=%~n0

:: This is the command the user is requesting.
set command=%1

:: Case statement for commands.
if "%command%"=="" goto help
if %command%==help goto help
if %command%==open goto open
if %command%==path goto path
if %command%==test goto test

:: Either help was called or command was not recognized.
:help
echo.
echo Description of %fileName%.
echo.
echo Usage: %fileName% ^<command^>
echo.
echo The commands are:
echo.
echo        test      runs a simple echo test
echo        open      opens the folderpath to this %fileName% script
echo        path      is the path to this %fileName% script
echo        help      is what you are currently looking at =^)
echo.
echo located at %path%
echo.
goto end

:: Various commands.
:open
start %path%
goto end

:path
echo.
echo %path%
goto end

:test
echo.
echo This is a test.
goto end

:: End of the call. Run any cleanup here if needed.
:end
