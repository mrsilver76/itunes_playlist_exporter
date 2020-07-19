# itunes_playlist_exporter
A script which connects to iTunes and exports all playlists in m3u format. It can also (optionally) adjust the paths of playlists to support NAS drives and upload them to Plex.

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

1.  Download the latest version from https://github.com/mrsilver76/itunes_playlist_exporter/releases
2.  Extract the files and place the file `iTunes_Playlist_Exporter.vbs` somewhere easy to reference when you want to run it
3.  Edit `iTunes_Playlist_Exporter.vbs` using your favourite text editor (I recommend Notepad++ but Notepad will do)
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

### PATH_FIND and PATH_REPLACE

If you plan on allowing other users/machines/software to access the playlist then you may find that the location embedded in the playlist file isn't accessible for them. For example, `D:\My Music` won't be accessible by a NAS or a different computer. To solve this you can use `PATH_FIND` and `PATH_REPLACE` to swap out the paths within the playlists to ones which can be accessed. 

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

### PLEX_PLAYLIST_LOCATION

The full path required by Plex to access the playlists stored in `LOCAL_PLAYLIST_LOCATION`. If you're running Plex on a different machine to the one running the script, then the path that Plex needs to use to upload the playlists will be different. If you don't set this correctly then Plex cannot find the playlists.

## Command line options

    Usage: cscript.exe iTunes_Playlist_Exporter.vbs [/E] [/U]
    
        No args     Start export from iTunes and upload to Plex.
        /?          Display help.
        /E          Don't export from iTunes.
        /D          Don't delete playlists from Plex.
        /U          Don't upload to Plex. Implies /E.

Options `/E`, `/D` and `/U` are useful when you're trying to upload multiple playlists to Plex. For more details see the later section.

### Don't export from iTunes (/E)

Skips the deleting of previous playlists and exporting of new playlists from iTunes.

### Don't delete playlists from Plex (/D)

Skips the deleting of all playlists on Plex prior to uploading new ones.

#### Don't upload to Plex (/U)

Skips the deleting of existing playlists on Plex and the re-uploading of the new playlists.

## Common usage scenarios

### Single machine with iTunes export

You want to export all your playlists from iTunes into a folder. You plan to only use the playlists on this machine.

*  Configure `LOCAL_PLAYLIST_LOCATION` and `IGNORE_PREFIX`
*  Run the script (`/D` is optional as Plex isn't configured correctly)

### iTunes export to networked drive for use with music software running on that networked drive can access

You want to export all your playlists from iTunes into a folder on a network (eg. NAS) so that software on that same networked drive can access it.

*  Configure `LOCAL_PLAYLIST_LOCATION` and `IGNORE_PREFIX`
*  Configure `PATH_FIND` to be the location of the music on your computer
*  Configure `PATH_REPLACE` to be the location of the music as your network drive would see it
*  Run the script 

### iTunes export to Plex running on same machine

You have iTunes and Plex running on the same machine.

*  Configure `LOCAL_PLAYLIST_LOCATION` and `IGNORE_PREFIX`
*  Configure `TOKEN`, `SERVER` and `LIBRARY_ID`
*  Configure `PLEX_PLAYLIST_LOCATION` to be the same as `LOCAL_PLAYLIST_LOCATION`
*  Run the script 

### iTunes export to Plex running on different machine

You have Plex running on a different machine to iTunes. Your playlists need to be copied to somewhere that both Plex and your PC running iTunes can access. There are three possible locations for the playlists:

1.  Stored on the machine with iTunes, but with a network share so that it can be accessed by other machines
2.  Stored on the machine with Plex, but with a network share so that it can be accessed by other machines
3.  Stored on a different machine (eg. NAS), with a network path that can be accessed by all other machines

*  Configure `LOCAL_PLAYLIST_LOCATION` to be network path to where the playlists should be stored. This location needs to be accessible by all computers.
*  Configure `IGNORE_PREFIX`
*  Configure `TOKEN`, `SERVER` and `LIBRARY_ID`
*  Configure `PLEX_PLAYLIST_LOCATION` to be the location where Plex would look to find the playlists
*  Run the script 

#### Multiple iTunes exports

See the next section.

## Managing multiple playlists from multiple users with Plex

This section caters for people who run multiple computers with iTunes and wish to store all their music and playlists in a single location (eg. a NAS - but could be a computer with a shared 
