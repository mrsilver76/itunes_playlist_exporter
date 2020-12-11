# iTunes Playlist exporter
A script which connects to iTunes and exports all playlists in m3u format. It can also (optionally) adjust the paths of playlists to support remote drives (eg. on a NAS or a shared drive on another computer) and upload them to Plex.

> :warning: **This script can delete playlists previously stored in Plex**: See the "Plex warning" section for more details.

## Features

This script has the following features:

1. Export the playlists to any folder, including network drives
2. Define playlists to be ignored based on their prefix
3. Optionally write out the playlists using Linux file path convention and with Linux newline support
4. Replace paths in a playlist with ones that can be accessed on by other users/machines/software (for example, repoint paths to a NAS)
5. Delete existing playlists on Plex and upload new ones
6. Command line options to prevent exporting of playlists and uploading of playlists (useful if trying to merge multiple playlists)

## Plex warning

This script allows you to (optionally) upload your playlists to Plex.

In order to keep playlists between iTunes and Plex in sync, the script treats iTunes as the primary store of playlists and mirrors the content to Plex. As a result _**all non-smart Plex playlists solely associated with a music library (defined within the script) will be deleted and replaced**_.

The following playlist types will not be deleted:
*  Playlists that contains content outside of the music library (eg. someone elses playlists or a playlist of video files)
*  Smart playlists generated within Plex 

> **If you do not want to delete playlists in Plex, then you should use `/D` from the command line.**

## Requirements

In order to use this program you need the following:

*   A computer running Microsoft Windows and iTunes
*   Some music stored in iTunes using a DRM free format (eg. MP3 or AAC)
*   One or more playlists for exporting
*   Knowledge of running command line applications

This program is not recommended for people who are not comfortable with the workings of Microsoft Windows and command line applications.

## Installation and usage

1.  Download the latest version from https://github.com/mrsilver76/itunes_playlist_exporter/releases
2.  Extract the files and place the file `iTunes_Playlist_Exporter.vbs` somewhere easy to reference when you want to run it
3.  Edit `iTunes_Playlist_Exporter.vbs` using your favourite text editor (I recommend [Notepad++](https://notepad-plus-plus.org/) but Notepad will do)
4.  Modify the settings at the top of the script based on the instructions below and then save it.
5.  Run the command by either double-clicking on the icon (a pop-up window will appear) or from the command line using: `cscript itunes_playlist_export.vbs`.
6.  There are some options you can call from the command line to change the use of the script. For more details see later.


## Script options

The following options need to be modified directly within the script.

### LOCAL_PLAYLIST_LOCATION

This is the full path to the folder you wish to save playlists to. If a folder doesn't exist then one will be created for you. If you plan on allowing other users/machines/software (such as Plex) to access these playlists then you need to pick a location which is either on a network drive or shared

If this is left blank then a folder will be created at `C:\Users\[your username]\Music\Exported iTunes Playlists` and used.

### IGNORE_PREFIX

This tells the script not to export any playlists that start with the following text. This is useful if you use playlists to help create other playlists and you don't
want these 'interim' ones being exported. 

For example, with a prefix of "x " then a playlist called "x Temporary" won't be exported.

### USE_LINUX_PATHS

This tells the script that the playlists will be accessed by a machine running Linux. One example of this would be if you intend on using the music software running on your NAS to access the playlists.

If you set this option to `True` then then following things will happen:

1.  All existing paths will have any `\` replaced with `/`
2.  The created playlist will use Linux line-endings (`\n`) instead of DOS line-endings (`\r\n`)

If you set this option to `False` then the paths and the playlist file format will be best suitable for Windows machines.

### PATH_FIND and PATH_REPLACE

If you plan on allowing other users/machines/software to access the playlist then you may find that the location embedded in the playlist file can only be accessed by the machine you created the playlist on. For example, `D:\My Music` won't be accessible by a NAS or a different computer. To solve this you can use `PATH_FIND` and `PATH_REPLACE` to swap out the paths within the playlists to ones which can be accessed. 

`PATH_FIND` - the text to find within the path of a song

`PATH_REPLACE` - the text to replace within the path of a song

## Plex specific options

These options relate to automatically uploading your playlists to Plex Media Server. Plex is free software (with an optional paid for pass) that allows you to manage and view your audio and video content. More details can be found at https://www.plex.tv/

When this script launches it will test connectivity to Plex. If this fails then the script will continue to run, but modifications to Plex will be disabled.

>**Note:** There is a bug within the Plex API which means that characters with [acute](https://en.wikipedia.org/wiki/Acute_accent) and [grave](https://en.wikipedia.org/wiki/Grave_accent) accents may appear as `?` within the Plex interface.

### TOKEN

The API token required for this script to be able to access Plex. For details on how to find this, see https://support.plex.tv/articles/204059436-finding-an-authentication-token-x-plex-token/.

Do not share your token with anyone!

### SERVER

The URL path to the Plex server and port number. This can start `http` or `https` depending on whether or not you're forcing secure connections. You can use 127.0.0.1 to mean "the same machine". Unless you've changed it, the port for Plex is  `32400`.

### LIBRARY_ID

The number of the library that the playlists will be loaded into. Make sure this library exists, is set up for music and contains all the songs that you have in the playlists. To find your library, go into the Plex web client, hover the mouse over the library you want and look at the URL. It will end with `source=xx` where `xx` is the library ID.

> :warning: **This script will delete any (non-smart) playlists that contains content solely from this LIBRARY_ID**


### PLEX_PLAYLIST_LOCATION

The full path required by Plex to access the playlists stored in `LOCAL_PLAYLIST_LOCATION`. If you're running Plex on a different machine to the one running the script, then the path that Plex needs to use to upload the playlists will be different. If you don't set this correctly then Plex cannot find the playlists.

## Command line options

    Usage: cscript.exe iTunes_Playlist_Exporter.vbs [/E] [/U]
    
        No args     Start export from iTunes and upload to Plex.
        /?          Display help.
        /E          Don't export from iTunes.
        /D          Don't delete playlists from Plex.
        /U          Don't upload to Plex.

### Don't export from iTunes (/E)

Skips the deleting of previous playlists and exporting of new playlists from iTunes.

### Don't delete playlists from Plex (/D)

Skips the deleting of all playlists on Plex prior to uploading new ones.

### Don't upload to Plex (/U)

Skips the deleting of existing playlists on Plex and the re-uploading of the new playlists.

## Uploading music and playlists to a shared drive

`Update_Music.bat` is a sample batch file which, when run, will do the following:

1. Copy all the music content from the local iTunes Media folder to a shared network drive (`\\OurNAS\Music\Songs` but can be changed)
2. Export all playlists to a shared network drive (`\\OurNAS\Music\Playlists` but can be changed)
3. Optionally upload to Plex (if `itunes_playlist_exporter` has been configured to do so)

You will need to edit the content of the batch file before it can be run. It's recommended that you place this batch file and the script on your shared network location and then create a shortcut on your desktop to it.

### Managing multiple playlists from multiple users to a shared drive

The `Update_Music.bat` batch file can support multiple users all uploading music and playlists to a shared drive.

To do this, it takes the Windows username of the person running the script (henceforce known as `Username`), copies music to `\\OurNAS\Music\Songs\Username` and copies playlists to `\\OurNAS\Music\Playlists\Username`. It will then attempt to launch `iTunes_Playlist_Exporter_Username.vbs` which is a copy of the script specifically configured for that user. You need to have a copy of the script (with the correct filename appended to it) for every user who plans to run the script.

You will need to edit the content of the batch file before it can be run. It's recommended that you place this batch file and the copies of the scripts (all with different usernames appended) on your shared network location and then create a shortcut on each persons desktop to it.

## Questions, comments or suggestions?

The easiest way is to raise a ticket on Github.
