# itunes_playlist_exporter
A script which connects to iTunes and exports all playlists in m3u format. It can also (optionally) adjust the paths of playlists to support remote drives (eg. on a NAS or a shared drive on another computer) and upload them to Plex.

> :warning: **This script will delete playlists previously stored in Plex**: See the "warning" section for more details.

## Features

This script has the following features:

1. Export the playlists to any folder, including network drives
2. Define playlists to be ignored based on their prefix
3. Replace paths in a playlist with ones that can be accessed on by other users/machines/software (for example, repoint paths to a NAS)
4. Delete existing playlists on Plex and upload new ones
5. Command line options to prevent exporting of playlists and uploading of playlists (useful if trying to merge multiple playlists)

## Requirements

In order to use this program you need the following:

*   A computer running Microsoft Windows and iTunes
*   Some music stored in iTunes using a DRM free format (eg. MP3 or AAC)
*   One or more playlists for exporting
*   Knowledge of running command line applications.

This program is not recommended for people who are not comfortable with the workings of Microsoft Windows and command line applications.

## Installation and usage

1.  Download the latest version from https://github.com/mrsilver76/itunes_playlist_exporter/archive/master.zip
2.  Extract the files and place the file `iTunes_Playlist_Exporter.vbs` somewhere easy to reference when you want to run it
3.  Edit `iTunes_Playlist_Exporter.vbs` using your favourite text editor (I recommend Notepad++ but Notepad will do)
4.  Modify the settings at the top of the script based on the instructions below and then save it.
5.  Run the command by either double-clicking on the icon (a pop-up window will appear) or from the command line using: `cscript itunes_playlist_export.vbs`.
6.  There are some options you can call from the command line to change the use of the script. For more details see later.

## Known issue/restrictions

This script has a number of known issues/restrictions. If you are able to help, please add commentary in the appropriate ticket.

1.  Basic synchronisation is achieved by deleting all playlists from Plex and then re-uploading. This is because there is no Plex API call to just delete the playlists associated with a single library. (github issue)
2.  The titles of playlists in Plex do not match the original playlist name. This is because the playlist filename needs to remove all characters not supported by Windows but also a range of characters not supported by Plex. There does not appear to be an API call to rename a playlist that has been imported into Plex (github issue)
3.  Some playlists do not import into Plex and produce an internal error. There is no indication as to why this happens in the error message or Plex logs (github issue)
4.  iTunes music files which have DRM are not supported. There is no workaround for this.

> :warning: It is not recommended to use this script if you create and maintain playlists within Plex.

## Script options

The following options need to be modified directly within the script.

### LOCAL_PLAYLIST_LOCATION

This is the full path to the folder you wish to save playlists to. If a folder doesn't exist then one will be created for you. If you plan on allowing other users/machines/software (such as Plex) to access these playlists then you need to pick a location which is either on a network drive or shared

If this is left blank then a folder will be created at `C:\Users\[your username]\Music\Exported iTunes Playlists` and used.

### IGNORE_PREFIX

This tells the script not to export any playlists that start with the following text. This is useful if you use playlists to help create other playlists and you don't
want these 'interim' ones being exported. 

For example, with a prefix of "x " then a playlist called "x Temporary" won't be exported.

### PATH_FIND and PATH_REPLACE

If you plan on allowing other users/machines/software to access the playlist then you may find that the location embedded in the playlist file can only be accessed by the machine you created the playlist on. For example, `D:\My Music` won't be accessible by a NAS or a different computer. To solve this you can use `PATH_FIND` and `PATH_REPLACE` to swap out the paths within the playlists to ones which can be accessed. 

`PATH_FIND` - the text to find within the path of a song

`PATH_REPLACE` - the text to replace within the path of a song

## Plex specific options

These options relate to automatically uploading your playlists to Plex Media Server. Plex is free software (with an optional paid for pass) that allows you to manage and view your audio and video content. More details can be found at https://www.plex.tv/

### TOKEN

The API token required for this script to be able to access Plex. For details on how to find this, see https://support.plex.tv/articles/204059436-finding-an-authentication-token-x-plex-token/.

Do not share your token with anyone!

### SERVER

The URL path to the Plex server and port number. This can start `http` or `https` depending on whether or not you're forcing secure connections. You can use 127.0.0.1 to mean "the same machine". Unless you've changed it, the port for Plex is  `32400`.

### LIBRARY_ID

The number of the library that the playlists will be loaded into. Make sure this library exists, is set up for music and contains all the songs that you have in the playlists. To find your library, go into the Plex web client, hover the mouse over the library you want and look at the URL. It will end with `source=xx` where `xx` is the library ID.

> :warning: **This script will delete playlists previously stored in Plex!**: If you have created playlists solely in Plex then they will be lost.


### PLEX_PLAYLIST_LOCATION

The full path required by Plex to access the playlists stored in `LOCAL_PLAYLIST_LOCATION`. If you're running Plex on a different machine to the one running the script, then the path that Plex needs to use to upload the playlists will be different. If you don't set this correctly then Plex cannot find the playlists.

## Command line options

    Usage: cscript.exe iTunes_Playlist_Exporter.vbs [/E] [/U]
    
        No args     Start export from iTunes and upload to Plex.
        /?          Display help.
        /E          Don't export from iTunes.
        /D          Don't delete playlists from Plex.
        /U          Don't upload to Plex.

Options `/E` and `/D` are useful when you're trying to upload multiple playlists to Plex. For more details see the later section.

### Don't export from iTunes (/E)

Skips the deleting of previous playlists and exporting of new playlists from iTunes.

### Don't delete playlists from Plex (/D)

Skips the deleting of all playlists on Plex prior to uploading new ones.

### Don't upload to Plex (/U)

Skips the deleting of existing playlists on Plex and the re-uploading of the new playlists.

## Managing multiple playlists from multiple users to a shared drive (and, optionally, uploading to Plex)

This section will be useful for people who run multiple Windows computers with iTunes and wish to store all their music and playlists in a single location such as a NAS. It doesn't have to be a NAS, any computer with a shared folder could be used.

In our example, we're going to use three people (`Rita`, `Bob` and `Sue`) who all have a laptop running iTunes and a network drive (called `OurNAS`). We also assume that "Keep iTunes Media folder organised" and "Copy files to iTunes Media folder when adding to library" are  enabled on everyones laptops.

### Copying the music files to the network drive

The easiest way to do this is to use robocopy and mirror the contents of the iTunes music folder to the network drive. We'll create a single script which will use the username to determine where the files should be located. The contents will go in a folder called "Music" which, itself, is seperated into "Songs" and "Playlists"

    @echo off
    setlocal enabledelayedexpansion
    
    SET src=C:\Users\%username%\Music\iTunes\iTunes Media\Music
    SET music=\\OurNAS\Music\Songs\%username%
    SET playlists=\\OurNAS\Music\Playlists\%username%

    robocopy /MIR /COPY:DAT /DCOPY:DAT /MT /R:5 /W:5 "%src%" "%music%"

Now we need to export the playlists:

    cscript "iTunes_Playlist_Exporter_%username%.vbs"
    
This will execute one of three scripts (`iTunes_Playlist_Exporter_Rita.vbs`, `iTunes_Playlist_Exporter_Bob.vbs` or `iTunes_Playlist_Exporter_Sue.vbs`). We now need to make three copies of the iTunes Playlist script and modify each of them so that they work for that user on their machine. In other words, if Rita runs `iTunes_Playlist_Exporter_Rita.vbs` then it should export her playlists to `\\OurNAS\Music\Playlists\Rita` - like so:

    Const LOCAL_PLAYLIST_LOCATION = "\\OurNAS\Music\Playlists\Rita"
    Const PATH_FIND = "C:\Users\Rita\Music\iTunes\iTunes Media\Music\"
    Const PATH_REPLACE = "\\OurNAS\Music\Songs\Rita\"

After each person has run the script once we should have:

1.  `\\OurNAS\Music\Songs` containing three folders (`Rita`, `Bob` and `Sue`) with music in each of them
2.  `\\OurNAS\Music\Playlists` also containing three folders (`Rita`, `Bob` and `Sue`) with playlists in each of them
3.  Each of the playlists should correctly reference the location of the song on the NAS, rather than each persons own computer

If you're using software on the NAS (or shared network drive) which isn't Plex, then you should now configure it to find the playlists and the songs.

### Uploading to Plex

If you're using Plex then you cannot point it to a folder of m3u playlists, the contents must be uploaded to the Plex server through the API.

If we just call this script three times (one for Rita, then Bob and finally Sue) then we'll have two problems:

1.  Two out of the three scripts will fail to export from iTunes because they aren't being run on the correct computer
2.  Although uploading of the playlists to Plex will work, the existing Plex playlists will be erased each time it is run - leaving you with only Sue's.

To resolve this we need to export and upload the main users first:

    cscript "iTunes_Playlist_Exporter_%username%.vbs"

and then export everyone elses playlists using the /E and /D flags to tell the script not to export from iTunes (which will fail) and not to delete the playlists which are already loaded into Plex:

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

This script should be stored on a network share with a shortcut placed on Rita, Bob and Sue's computers.

A version of the script with comments can be found here.

## Questions, comments or suggestions?

The easiest way is to raise a ticket on Github.
