@echo off
setlocal enabledelayedexpansion

rem -=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
rem
rem Demonstration batch file which will copy all songs and playlists
rem stored in an iTunes folder to a network share.
rem
rem Once uploaded, you can use your favourite music software to access
rem the songs and playlists.
rem
rem -=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=

rem Set the variables
rem
rem   src        = The location of the iTunes Music folder on the computer.
rem   music      = The location on the network share where the music library
rem                should be copied to.
rem   playlists  = The location on the network share where the playlists should be
rem                copied to. 
rem
rem  Notes: * Music and playlists are copied to a network drive to ensure that
rem           the content is available even when the computer is running
rem         * %username% is the username of the person running this script

SET src=C:\Users\%username%\Music\iTunes\iTunes Media\Music
SET music=\\OurNAS\Music\Songs
SET playlists=\\OurNAS\Music\Playlists

rem Mirror the contents of the iTunes Media folder on the computer to the network
rem share. Uses multiple threads for speed and will retry up to five times in case
rem the file is locked by another program or virus scanner.

robocopy /MIR /COPY:DAT /DCOPY:DAT /MT /R:5 /W:5 "%src%" "%music%"

rem Export the playlists from iTunes to the playlists location defined within the script.
rem
rem If there is a verison of the script with the current username appended to the end
rem of the filename (eg. iTunes_Playlist_Exporter_Bob.vbs") then use that instead for
rem multi-user support.

if exist "iTunes_Playlist_Exporter_%username%.vbs" (
	cscript /nologo "iTunes_Playlist_Exporter_%username%.vbs"
) else (
	cscript /nologo "iTunes_Playlist_Exporter.vbs"
)

echo "Finished"
timeout 5