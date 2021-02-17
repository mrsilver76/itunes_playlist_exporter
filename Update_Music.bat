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

rem Are we trying to support multiple users? If there is a version of this
rem script but with a username appended then the answer is "yes"

if exist "iTunes_Playlist_Exporter_%username%.vbs" (
	set multiuser=1
) else (
	set multiuser=0
)

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

rem Here we are assuming that everyone stores their music in the same location
rem but with a slightly different username. If this is not the case then you'll
rem need to use "if" statements to configure manually. For example:
rem
rem   if "%username%" == "Rita" (
rem   	SET src=C:\Users\Rita\Music\iTunes\iTunes Music\Music
rem   )
rem   if "%username%" == "Bob" (
rem   	SET src=D:\Music\iTunes
rem   )

SET src=C:\Users\%username%\Music\iTunes\iTunes Media\Music

rem This is the location on the shared drive for the content to be copied to.
rem %username% is only used if we are doing multi-user.

if "%multiuser%" == "1" (
	rem Multiple users, so we need a folder for each user
	SET music=\\OurNAS\Music\Songs\%username%\
	SET playlists=\\OurNAS\Music\Playlists\%username%\
) else (
	rem Single user, so just put everyting in one folder
	SET music=\\OurNAS\Music\Songs\
	SET playlists=\\OurNAS\Music\Playlists\
)

rem Nothing to change below here
rem -=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=

rem Mirror the contents of the iTunes Media folder on the computer to the network
rem share. Uses multiple threads for speed and will retry up to five times in case
rem the file is locked by another program or virus scanner.

robocopy /MIR /FFT /COPY:DAT /DCOPY:DAT /MT /R:5 /W:5 "%src%" "%music%"

rem Export the playlists from iTunes to the playlists location defined within the script.
rem
rem If there is a verison of the script with the current username appended to the end
rem of the filename (eg. iTunes_Playlist_Exporter_Bob.vbs") then use that instead for
rem multi-user support.

if "%multiuser%" == "1" (
	cscript /nologo "iTunes_Playlist_Exporter_%username%.vbs"
) else (
	cscript /nologo "iTunes_Playlist_Exporter.vbs"
)

echo "Finished"
timeout 5