@echo off
setlocal enabledelayedexpansion

rem -=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
rem
rem Demonstration batch file which will copy all songs and playlists
rem from three computers (owned by Rita, Bob and Sue) to a network share
rem (called OurNAS).
rem
rem -=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=

rem Set the variables
rem
rem   src        = The location of the iTunes Music folder on the computer. We're
rem                assuming that all three computers store the folder in the same
rem                place with only the username different.
rem   music      = The location on the network share where each persons music library
rem                should be copied to. This is to ensure that the music is available
rem                even when the computer is turned off.
rem   playlists  = The location on the network share where the playlists should be
rem                copied to. This is to ensure that the playlists are available even
rem                when the computer is turned off.
rem
rem  Note: %username% is the username of the person running this script

SET src=C:\Users\%username%\Music\iTunes\iTunes Media\Music
SET music=\\OurNAS\Music\Songs\%username%
SET playlists=\\OurNAS\Music\Playlists\%username%

rem Mirror the contents of the iTunes Media folder on the computer to the network
rem share. Uses multiple threads for speed and will retry up to five times in case
rem the file is locked by another program or virus scanner.

robocopy /MIR /COPY:DAT /DCOPY:DAT /MT /R:5 /W:5 "%src%" "%music%"

rem Export the playlists from iTunes to the playlists location defined within the script.
rem Note that each person should have their own version of the exporter script which
rem has been configured to correctly run on their computer.

cscript "iTunes_Playlist_Exporter_%username%.vbs"

rem -=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
rem
rem Optional section to upload outstanding playlists to Plex
rem
rem -=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=

rem We've just deleted all playlists on Plex and uploaded updated playlists
rem for the person running this script. We now need to upload everyone elses
rem playlists (which are currently stored on the network drive).
rem
rem Note that we use /E to prevent the script from trying to export (which will
rem fail, as it's the wrong computer) and /D to prevent the script from deleting
rem the playlists we have.

if "%username%" == "Rita" (
    cscript "iTunes_Playlist_Exporter_Bob.vbs" /E /D
    cscript "iTunes_Playlist_Exporter_Sue.vbs" /E /D
) 
if "%username%" == "Bob" (
    cscript "iTunes_Playlist_Exporter_Rita.vbs" /E /D
    cscript "iTunes_Playlist_Exporter_Sue.vbs" /E /D
) 
if "%username%" == "Sue" (
    cscript "iTunes_Playlist_Exporter_Rita.vbs" /E /D
    cscript "iTunes_Playlist_Exporter_Bob.vbs" /E /D
) 