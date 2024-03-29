Option Explicit

' iTunes Playlist Exporter v1.7.0, Copyright © 2020-2022 Richard Lawrence
' https://github.com/mrsilver76/itunes_playlist_exporter
'
' A script which connects to iTunes and exports all playlists in m3u
' format. It can also (optionally) adjust the paths of playlists to
' support NAS drives and upload them to Plex.
'
' Notes: * DRM'ed music is not exported and not supported by Plex.
'        * Existing playlists created (with only content from the library
'          you're uploading to) will be deleted.
'        * If you're not running Windows 10, you need to install Curl which
'          is available from https://curl.haxx.se/windows/
'
' ----- Licence ----------------------------------------------------------
'
' This program is free software; you can redistribute it and/or
' modify it under the terms of the GNU General Public License
' as published by the Free Software Foundation; either version 2
' of the License, or (at your option) any later version.
'
' This program is distributed in the hope that it will be useful,
' but WITHOUT ANY WARRANTY; without even the implied warranty of
' MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
' GNU General Public License for more details.
' 
'
' ----- Configuration settings -------------------------------------------

' LOCAL_PLAYLIST_LOCATION - the location to store the generated playlists. If
' you plan on allowing other users/machines/software (such as Plex) to access
' these playlists then you need to pick a location which is either on a network
' drive or shared. If this is left blank then a folder will be created at 
' "C:\Users\username\Music\Exported iTunes Playlists" and used.

Const LOCAL_PLAYLIST_LOCATION = ""

' IGNORE_PREFIX - don't export any playlists that start with the following text.
' This is useful if you use playlists to help create other playlists and you don't
' want these 'interim' ones being exported. 

Const IGNORE_PREFIX = "z "

' If you plan on allowing other users/machines/software to access the playlist
' then you may find that the location embedded in the playlist file isn't accessible
' for them. For example, "D:\MyMusic" won't be accessible by a NAS or another
' computer. To solve this you can use the following options to modify your paths
' so they point to the correct location.

' USE_LINUX_PATHS - if set to True will cause two things to happen:
'                     1. All existing paths will replace \ with /
'                     2. The created playlist will use Linux line-endings (\n)
'                        instead of DOS line-endings (\r\n)

Const USE_LINUX_PATHS = False

' These two options allow you to search (PATH_FIND) for a string in a path
' and then replace (PATH_REPLACE) it with something else.
'
' PATH_FILE - the text to find within the path of a song

Const PATH_FIND = "C:\Users\Richard\Music\iTunes\iTunes Music\Music\"

' PATH_REPLACE - the test to replace within the path of a song

Const PATH_REPLACE = "\\Storage\Content\Music\Songs\Richard\"

' ----- Plex specific configuration settings ---------------------

' These options relate to automatically uploading your playlists to Plex
' Media Server. Plex is free software (with an optional paid for pass) that
' allows you to manage and view your audio and video content. More details
' can be found at https://www.plex.tv/

' Note: Uploading will only happen if you provide a valid URL and token.
'       The script will happily run without these being provided.

' TOKEN - the API token required for this script to be able to access Plex.
' Do not share your token with anyone. For details on how to find this, see
' https://support.plex.tv/articles/204059436-finding-an-authentication-token-x-plex-token/

Const TOKEN=""

' SERVER - the URL path to the server and port number. This can start http
' or https depending on whether or not you're forcing secure connections. You
' can use 127.0.0.1 to mean "the same machine".

Const SERVER = "http://127.0.0.1:32400/"

' LIBRARY_ID - the number of the library that the playlists will be loaded into.
' Make sure this library exists, is set up for music and contains all the songs
' that you have in the playlists. To find your library, go into the Plex web client,
' hover the mouse over the library you want and look at the URL. It will end with
' "source=xx" where xx is the library ID.

Const LIBRARY_ID = 12

' PLEX_PLAYLIST_LOCATION - the full path required by Plex to access the playlists
' stored in LOCAL_PLAYLIST_LOCATION. If Plex is on a different machine to the 
' one running the script, this will have a different path.

Const PLEX_PLAYLIST_LOCATION = "\\Storage\Content\Music\Playlists\Richard"

' ----- End of configuration settings. Code starts here ------------------

Dim fso : Set fso = CreateObject("Scripting.FileSystemObject")
Dim wshShell : Set wshShell = CreateObject("WScript.Shell")

Dim sPlayListLocation : sPlayListLocation = Get_Playlist_Location(LOCAL_PLAYLIST_LOCATION)
Dim sPlexPlayListLocation : sPlexPlayListLocation = Get_Playlist_Location(PLEX_PLAYLIST_LOCATION)
Dim bExportFromItunes, bUploadToPlex, bDeletePlexPlaylists, bIgnoreSmartPlaylists, bVerbose
bExportFromItunes = True
bUploadToPlex = True
bDeletePlexPlaylists = True
bIgnoreSmartPlaylists = False
bVerbose = False

Const VERSION = "1.7.0"

Call Force_Cscript_Execution

WScript.Echo "iTunes Playlist Exporter v" & VERSION & ", Copyright " & Chr(169) & " 2020-2022 Richard Lawrence"
WScript.Echo "https://github.com/mrsilver76/itunes_playlist_exporter"
WScript.Echo "This program comes with ABSOLUTELY NO WARRANTY. This is free software,"
WScript.Echo "and you are welcome to redistribute it under certain conditions; see"
WScript.Echo "the documentation for details."
WScript.Echo

Call Read_Params

Call Log("Starting iTunes Playlist Exporter")

If bVerbose = True Then Call Log("Verbose mode enabled")

' Don't upload to Plex if we cannot find the server
If Test_Plex() = False Then bUploadToPlex = False

If bExportFromItunes = True Then
	Call Delete_Existing_Playlists
	Call Export_Playlists
Else
	Call Log("Skipping exporting of playlists from iTunes")
End If

If bUploadToPlex = True Then
	If bDeletePlexPlaylists = True Then
		Call Delete_Playlists_From_Plex
	Else
		Call Log("Skipping deleting of playlists on Plex")
	End If
	Call Upload_Playlists
Else
	Call Log("Skipping upload of playlists to Plex")
End If

Call Log("Finished")
WScript.Quit


' -=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=

' Delete_Existing_Playlists
' Delete all playlists previously stored

Sub Delete_Existing_Playlists

	If fso.FolderExists(sPlayListLocation) = False Then
		Call Log("Creating folder for playlists in " & sPlaylistLocation)
		fso.CreateFolder(sPlayListLocation)
		Exit Sub
	End If

	' Folder exists, so just delete the contents
	Call Log("Deleting previous playlists in " & sPlaylistLocation)
	
	On Error Resume Next
	fso.DeleteFile(sPlayListLocation & "\*.m3u")
	On Error Goto 0
	
End Sub

' Export_Playlists
' Scan each playlist and write out m3u

Sub Export_Playlists

	Call Log("Connecting to iTunes to export playlists...")

	Dim oiTunes : Set oiTunes = CreateObject("iTunes.Application")
	Dim oMainLibrary : Set oMainLibrary = oiTunes.LibraryPlaylist
	Dim oTxt : Set oTxt = CreateObject("ADODB.Stream")
	Dim sContent

	If USE_LINUX_PATHS = True Then Call Log("Exported playlists will be in Linux format")

	Dim oPlayList : For Each oPlayList In oiTunes.LibrarySource.Playlists
		' Get the content of the playlist
		sContent = Get_Tracks(oPlayList)
		' If we have something then save it
		If sContent <> "" Then
			With oTxt
				.CharSet = "utf-8"
				.Open
				.WriteText sContent
				.SaveToFile sPlayListLocation & "\" & Clean_String(oPlayList.Name) & ".m3u", 2
				.Close
			End With
		End If
	Next

	Set oTxt = Nothing
	Set oiTunes = Nothing

End Sub

' Get_Tracks
' Get the playlists from iTunes and write them to a file

Function Get_Tracks(oPlayList)

	Dim sPath

	Get_Tracks = ""

	' Needs to be visible
	If oPlayList.Visible = False Then Exit Function

	' ... and not special
	If CLng(oPlayList.SpecialKind) <> 0 Then Exit Function
	
	' ... and not empty
	If CLng(oPlayList.Tracks.Count) = 0 Then
		Call Log("Empty playlist ignored: " & oPlayList.Name)
		Exit Function
	End If
	
	' ... and not a smart playlist (if we've asked to ignore them)
	If bIgnoreSmartPlaylists = True And oPlaylist.Smart = True Then
		Call Log("Smart playlist ignored: " & oPlayList.Name)
		Exit Function
	End If
	
	' ... and not prefixed by anything we don't want.
	If IGNORE_PREFIX <> "" and Len(oPlayList.Name) >= Len(IGNORE_PREFIX) Then
		If LCase(Left(oPlayList.Name, Len(IGNORE_PREFIX))) = LCase(IGNORE_PREFIX) Then 
			Call Log( "Matching prefix ignored: " & oPlayList.Name)
			Exit Function
		End If
	End If

	Dim iTotal : iTotal = oPlayList.Tracks.Count
	Dim iDone : iDone = 0
	Dim iLast : iLast = 999
	Dim sLocation, sPathFind, sPathReplace, sLineEnding
	
	' Use correct line ending
	If USE_LINUX_PATHS = True Then
		sLineEnding = VbLf
	Else
		sLineEnding = VbCrLf
	End If
	
	' Work out what we are finding and replacing
	If PATH_FIND <> "" Then
		sPathFind = PATH_FIND
		sPathReplace = PATH_REPLACE
		' If we're using Linux paths then make sure that we aren't still using / for finding
		' as it'll always fail
		If USE_LINUX_PATHS = True Then
			sPathFind = Replace(sPathFind, "\", "/")
			sPathReplace = Replace(sPathReplace, "\", "/")
		End If
	End If
		
	' Fake a log entry so that it can be updated every second
	WScript.StdOut.Write "[" & FormatDateTime(Now(), vbLongTime) & "] Exporting playlist: " & oPlayList.Name & "  [" & oPlayList.Tracks.Count & " song" & Pluralise(oPlayList.Tracks.Count) & " / 0% completed]" & VbCr
	
	Dim oTrack : For Each oTrack In oPlayList.Tracks

		' Display progress
		If Second(Now()) <> iLast Then
			WScript.StdOut.Write "[" & FormatDateTime(Now(), vbLongTime) & "] Exporting playlist: " & oPlayList.Name & "  [" & oPlayList.Tracks.Count & " song" & Pluralise(oPlayList.Tracks.Count) & " / " & Int((iDone*100)/iTotal) & "% completed]" & VbCr
			iLast = Second(Now())
		End If

		If oTrack.Kind = 1 Then
			Select Case LCase(fso.GetExtensionName(oTrack.Location))
				Case "mp3", "m4a"
					' Make sure we have the header
					If Get_Tracks = "" Then
						Get_Tracks = "#EXTM3U" & sLineEnding
						' Add the name of the playlist
						Get_Tracks = Get_Tracks & "#PLAYLIST:" & oPlaylist.Name & sLineEnding
					End If
					' Add the track details
					Get_Tracks = Get_Tracks & "#EXTINF:" & oTrack.Duration & "," & oTrack.Name & " - " & oTrack.Artist & sLineEnding
					' And the path to the file
					sLocation = oTrack.Location
					' Are we using Linux paths?
					If USE_LINUX_PATHS = True Then sLocation = Replace(sLocation, "\", "/")
					' Do we need to replace something?
					If sPathFind <> "" Then
						Get_Tracks = Get_Tracks & Replace(sLocation, sPathFind, sPathReplace, 1, -1, VbTextCompare)
					Else
						Get_Tracks = Get_Tracks & sLocation
					End If
					' Add correct newline
					Get_Tracks = Get_Tracks & sLineEnding
				Case Else
					' Unwanted filetype, so skip it
			End Select
		End If
		iDone = iDone + 1
	Next
	
	' Confirm completed
	WScript.StdOut.Write Space(68+Len(oPlayList.Name)) & VbCr
	Call Log("Playlist exported: " & oPlayList.Name & "  [" & oPlayList.Tracks.Count & " song" & Pluralise(oPlayList.Tracks.Count) & "]")
	
End Function

' Force_Cscript_Execution
' Force the script to run using cscript instead of wscript.

Sub Force_Cscript_Execution

	Dim Arg, Str
    If Not LCase(Right(WScript.FullName, 12)) = "\cscript.exe" Then
        For Each Arg In WScript.Arguments
			If InStr(Arg, " ") Then Arg = """" & Arg & """"
			Str = Str & " " & Arg
		Next
        CreateObject("WScript.Shell" ).Run "cscript //nologo """ & WScript.ScriptFullName & """ " & Str
        WScript.Quit
    End If
	
End Sub

' Pluralise
' Return a "s" or blank depending on if it's needed for plural words

Function Pluralise(iNum)

	If iNum = 1 Then
		Pluralise = ""
	Else
		Pluralise = "s"
	End If

End Function

' Clean_String
' Take a string, replace all accented characters with non-accented versions
' (to try and keep legibility) and then remove any characters which will
' prevent Windows or Plex from handling the filename.

Function Clean_String(sText)

	Dim objRegExp : Set objRegExp = New Regexp
	Dim sNew, iPos, iCode

	' Replace certain acented characters with their non-acented
	' equivilants
	
	For iPos = 1 To Len(sText)
		iCode = Asc(Mid(sText, iPos, 1))
		Select Case iCode
			Case 192, 193, 194, 195, 196, 197
				Clean_String = Clean_String & "A"
			Case 199
				Clean_String = Clean_String & "C"
			Case 200, 201, 202, 203
				Clean_String = Clean_String & "E"
			Case 204, 205, 206, 207
				Clean_String = Clean_String & "I"
			Case 208
				Clean_String = Clean_String & "D"
			Case 209
				Clean_String = Clean_String & "N"
			Case 210, 211, 212, 213, 214, 216
				Clean_String = Clean_String & "O"
			Case 217, 218, 219, 220
				Clean_String = Clean_String & "U"
			Case 221
				Clean_String = Clean_String & "Y"
			Case 232, 225, 226, 227, 228, 229
				Clean_String = Clean_String & "a"
			Case 232, 233, 234, 235
				Clean_String = Clean_String & "e"
			Case 236, 237, 238, 239
				Clean_String = Clean_String & "i"
			Case 240, 242, 243, 244, 245, 248
				Clean_String = Clean_String & "o"
			Case 241
				Clean_String = Clean_String & "n"
			Case 249, 250, 251, 252
				Clean_String = Clean_String & "u"
			Case 253, 255
				Clean_String = Clean_String & "y"				
			Case Else
				Clean_String = Clean_String & Chr(iCode)
		End Select
	Next
	
	' Clean_String now contains a cleaner title for the
	' m3u, but may still have characters in the filename which
	' will call Windows and Plex to fall over. So we need to remove these.

	With objRegExp
		.IgnoreCase = True
		.Global = True
		
		' Replace all invalid characters with spaces
		.Pattern = "[^a-z0-9 \-\)\(]"
		Clean_String = .Replace(Clean_String, "_")

		' Replace more than one space with a single space
		.Pattern = " +"
		Clean_String = .Replace(Clean_String, " ")
	End With
	
	Set objRegExp = Nothing

End Function

' Upload_Playlists
' Upload a m3u filename to Plex into the appropriate section

Sub Upload_Playlists

	Call Log("Uploading playlists to Plex at " & SERVER)

	Dim oFolder : Set oFolder = fso.GetFolder(sPlayListLocation)

	Dim oFile, sCmd, sOutput
	For Each oFile In oFolder.Files
		If LCase(fso.GetExtensionName(oFile.Name)) = "m3u" Then
		
			' Get the proper name of the playlist
			Dim sPlayListName : sPlayListName = Get_Playlist_Name(oFile.Path)
			
			' If the filename and the name of the playlist are the same then 
			' there is no point using a temporary file and then renaming - so
			' just upload it
			
			If fso.GetBaseName(oFile.Name) = sPlayListName Then
				Call Do_Upload_To_Plex(oFile.ParentFolder & "\" & oFile.Name, sPlaylistName)
			Else
				' We need to rename to something temporary, upload that file,
				' find the playlist with that temporary name and then
				' re-name it. All because Plex cannot upload a playlist with
				' a title!

				' Copy the file to a random unique name. Make sure that nothing else
				' is called that and then make a copy of the file to that name
				Dim sTempFile, sTempName
				Do
					sTempName = fso.GetBaseName(fso.GetTempName)
					sTempFile = oFile.ParentFolder & "\" & sTempName & ".m3u"
				Loop Until fso.FileExists(sTempFile) = False
				fso.CopyFile oFile.Path, sTempFile
			
				' Upload this to Plex
				Call Do_Upload_To_Plex(sTempFile, sPlaylistName)
							
				' Delete the temporary file
				On Error Resume Next
				fso.DeleteFile(sTempFile)
				On Error Goto 0
				
				' Find the unique ratingKey for this random unique name and then rename it
				Dim sTitle, sLine, sKey
				sOutput = Execute_Command("curl -sS """ & SERVER & "playlists/all/?X-Plex-Token=" & TOKEN & """", True)	
				For Each sLine In sOutput
					sKey = Find_From_Regexp(sLine, "ratingKey=\""(\d+?)\""")
					' Do we have a key and is it not a smart playlist
					If sKey <> "" And Instr(sLine, "smart=""0""") > 0 Then
						sTitle = Find_From_Regexp(sLine, "title=\""(.+?)\""")
						' Does the title of the playlist match our temporary name?
						If sTitle = sTempName Then
							' Rename it to the correct name
							Dim sRen: sRen = Execute_Command("curl -sS -X PUT """ & SERVER & "playlists/" & sKey & "/?title=" & URL_Encode(sPlaylistName) & "&X-Plex-Token=" & TOKEN & """", False)
							If sRen <> "" Then Call Log("Playlist renaming failed with: " & sRen)			
							Exit For
						End If
					End If
				Next
			End If
		End If
	Next

End Sub

' Do_Upload_To_Plex
' Takes a full path and filename of the upload along with a display name (in case it's
' being renamed at a later date) and actually does the upload to Plex

Sub Do_Upload_To_Plex(sUploadFile, sDisplayName)

	Dim sSlash
	If USE_LINUX_PATHS = True Then
		sSlash = "/"
	Else
		sSlash = "\"
	End If

	' Use the same filename as the upload if we haven't defined one
	if sDisplayName = "" Then sDisplayName = fso.GetBaseName(sUploadFile)

	Call Log("Adding playlist: " & sDisplayName)
	Dim sOutput : sOutput = Execute_Command("curl -sS -X POST """ & SERVER & "playlists/upload?sectionID=" & LIBRARY_ID & "&path=" & URL_Encode(sPlexPlaylistLocation & sSlash & fso.GetFileName(sUploadFile)) & "&X-Plex-Token=" & TOKEN & """", False)
	If sOutput <> "" Then Call Log("Curl failed with: " & sOutput)			
	
End Sub

' Delete_Playlists_From_Plex
' Delete all the playlists stored on Plex.
'
' Note: Calling /playlists/ with DELETE doesn't work, so we need to list
' all the ones which exist and then delete each one manually.

Sub Delete_Playlists_From_Plex

	Dim iTotal, iDone, iDeleted, iFailed, iPlaylists, iSmart, iSkipped, iAnalysed : iDone = 0 : iDeleted = 0 : iFailed = 0 : iPlaylists = 0 : iSmart = 0 : iSkipped = 0 : iAnalysed = 0

	Call Log("Deleting playlists associated with library ID " & LIBRARY_ID & " from " & SERVER)

	' Get a list of all the playlists for that library
	Dim sOutput : sOutput = Execute_Command("curl -sS """ & SERVER & "playlists/all/?X-Plex-Token=" & TOKEN & """", True)

	' Work out how many we have to process
	Dim sLine, sKey
	For Each sLine In sOutput
		sKey = Find_From_Regexp(sLine, "leafCount=\""(\d+?)\""")
		' Do we have a key and is it not a smart playlist and not a video playlist
		If sKey <> "" Then
			If Instr(sLine, "smart=""0""") > 0 And Instr(sLine, "playlistType=""audio""") > 0 Then
				iTotal = iTotal + CLng(sKey)
				iPlaylists = iPlaylists + 1
			Else
				iSmart = iSmart + 1
			End If
		End If
	Next

	Call Log("Found " & iPlaylists & " playlist" & Pluralise(iPlayLists) & " to analyse, containing " & iTotal & " song" & Pluralise(iTotal) & ". " & iSmart & " smart/video playlist" & Pluralise(iSmart) & " ignored")

	WScript.StdOut.Write "[" & FormatDateTime(Now(), vbLongTime) & "] 0% complete ... (0 analysed, 0 skipped, 0 deleted, 0 failed)" & VbCr
	
	Dim tStartTime : tStartTime = Now()

	' Now walk through the list again, extracting the ratingKey
	For Each sLine In sOutput
		sKey = Find_From_Regexp(sLine, "ratingKey=\""(\d+?)\""")
		' Only do something if we have a key and it's not a smart playlist and not a video playlist
		If sKey <> "" And Instr(sLine, "smart=""0""") > 0 And Instr(sLine, "playlistType=""audio""") > 0 Then	
			iAnalysed = iAnalysed + 1		
			' Verify if all of the items in this playlist can be deleted
			If All_Playlist_Contents_In_Library(sKey) = True Then			
				' Delete it from Plex
				Dim sDel : sDel = Execute_Command("curl -sS -X DELETE """ & SERVER & "playlists/" & sKey & "?X-Plex-Token=" & TOKEN & """", False)
				If sDel <> "" Then
					Call Log("Curl failed with: " & sDel)
					iFailed = iFailed + 1
				Else
					iDeleted = iDeleted + 1
				End If
			Else
				iSkipped = iSkipped + 1
			End If
			
			' Might as well re-use sKey again to find the number of playlists processed
			sKey = Find_From_Regexp(sLine, "leafCount=\""(\d+?)\""")
			iDone = iDone + CLng(sKey)
			
			WScript.StdOut.Write "[" & FormatDateTime(Now(), vbLongTime) & "] " & Int((iDone*100)/iTotal) & "% completed ... (" & iAnalysed & " analysed, " & iSkipped & " skipped, " & iDeleted & " deleted, " & iFailed & " failed)" & VbCr
					
		End If
	Next

	Call Log(iDeleted & " playlist" & Pluralise(iDeleted) & " deleted on Plex (with " & iFailed & " failure" & Pluralise(iFailed) & ")                         ")

End Sub

' All_Playlist_Contents_In_Library
' Given a playlist ID, looks at all the content sitting in that playlist. If there
' are any items which don't belong to the LIBRARY_ID then return False because this
' playlist should not be deleted. Else return True.

Function All_Playlist_Contents_In_Library(sKey)

	All_Playlist_Contents_In_Library = False
	
	' Get the playlist details
	Dim sOutput : sOutput = Execute_Command("curl -sS """ & SERVER & "playlists/" & sKey & "/items?X-Plex-Token=" & TOKEN & """", True)

	' Ideally we'd use a regexp but since the playlists can be massive, this
	' means it's extremely slow. So we're going to use string matching instead.
	
	Dim sLine, sString
	sString = "librarySectionID=""" & LIBRARY_ID & """"

	For Each sLine In sOutput
		' Check if we have a line which contains librarySectionID
		If Instr(sLine, "librarySectionID=") > 0 Then
			' Check the match again, but with the LIBRARY_ID included
			If Instr(sLine, sString) = 0 Then
				' It's not there, this is not a library we want to delete
				Exit Function
			End If
		End If
	Next
	
	' Made it all the way here, so good to delete
	All_Playlist_Contents_In_Library = True
		
End Function

' Find_From_Regexp
' Given a string and a RegExp, find the key from within it and return it. Needed to
' delete the playlist

Function Find_From_Regexp(sString, sPattern)

	Find_From_Regexp = ""

	Dim oRE : Set oRE = New RegExp
	With oRE
		.Global = False
		.IgnoreCase = True
		.Pattern = sPattern
	End With
	
	Dim oMatch : Set oMatch = oRE.Execute(sString)
	If oMatch.Count = 1 Then Find_From_Regexp = oMatch.Item(0).Submatches(0)
	
	Set oMatch = Nothing
	Set oRE = Nothing
	
End Function


' Execute_Command
' Run a command and return the output it created. If bReturnArray is set to True
' then the output will be returned as an array instead of a string. True is
' highly recommended for large amounts of data as creating a single string is
' very slow.

Function Execute_Command(sCmd, bReturnArray)

	If bReturnArray = True Then
		Execute_Command = Call_Command(sCmd)
	Else
		Dim sOutput : sOutput = Call_Command(sCmd)
		' Now convert to string
		Execute_Command = ""
		Dim i : For i = 0 To UBound(sOutput)
			If IsEmpty(sOutput(i)) = False Then Execute_Command = Execute_Command & sOutput(i)
		Next
	End If
		
End Function

' Call_Command
' The actual calling of the command and then returning an array of the
' output

Function Call_Command(sCmd)

	Dim lStart, lEnd, lRecv, i : i = 0

	If bVerbose = True Then
		If Len(TOKEN) > 0 Then
			Call Log("Executing: " & Replace(sCmd, TOKEN, "PLEXTOKEN"))
		Else
			Call Log("Executing: " & sCmd)
		End If
	End If

	lStart = Timer()

	Dim oExec : Set oExec = wshShell.Exec(sCmd)
	Do Until oExec.StdOut.AtEndOfStream
		ReDim Preserve sOutput(i)
		sOutput(i) = oExec.StdOut.ReadLine()
		lRecv = lRecv + Len(sOutput(i))
		i = i + 1
	Loop
	Set oExec = Nothing

	If i > 0 Then
		' The command produced output
		Call_Command = sOutput
	Else
		' No output was produced, so return an empty array
		ReDim sOutput(0)
		Call_Command = sOutput
	End If
	
	lEnd = Timer()
	
	If bVerbose = True Then
		Dim sUnit : sUnit = "bytes"
		If lRecv > 1024 Then
			lRecv = LRecv / 1024
			sUnit = "kilobyes"
			If lRecv > 1024 Then
				lRecv = LRecv / 1024
				sUnit = "megabytes"
			End If
		End If	
		Call Log("Received " & FormatNumber(lRecv, 2) & " " & sUnit & " in " & FormatNumber(lEnd - lStart, 2) & " seconds")
	End If

End Function

' URL_Encode
' Take a string and encode it so it can be passed as a URL

Function URL_Encode(sString)

	Dim iPos, iAsc, sTmp
	
	For iPos = 1 to Len(sString)
		iAsc = Asc(Mid(sString, iPos, 1))
		'If iAsc = 32 Then
		'	URL_Encode = URL_Encode & "+"
		'Else
		If (iAsc < 123 And iAsc > 96) Or (iAsc < 91 And iAsc > 64) Or (iAsc < 58 And iAsc> 47) Then
			URL_Encode = URL_Encode & Chr(iAsc)
		Else
			sTmp = Trim(Hex(iAsc))
			If iAsc < 16 Then
				URL_Encode = URL_Encode & "%0" & sTmp
			Else
				URL_Encode = URL_Encode & "%" & sTmp
			End If
		End If
	Next
			
End Function

' Get_Playlist_Location
' Return a default location for the playlist if one isn't already specified.
' If one is, make sure that it doesn't have a trailing \

Function Get_Playlist_Location(sLocal)

	If sLocal = "" Then
		' Configuration is empty, so we need to return the default
		Get_Playlist_Location = wshShell.ExpandEnvironmentStrings("%HOMEDRIVE%") & wshShell.ExpandEnvironmentStrings("%HOMEPATH%") & "\Music\Exported iTunes Playlists"
		Exit Function
	End If

	' If there is a trailing \ then remove it
	If Right(sLocal, 1) = "\" Then
		Get_Playlist_Location = Left(sLocal, Len(sLocal)-1)
	Else
		Get_Playlist_Location = sLocal
	End If

End Function

' Read_Params
' Determine any command line options

Sub Read_Params

	Dim iCount : For iCount = 0 to WScript.Arguments.Count - 1
		Select Case LCase(WScript.Arguments(iCount))
			Case "/?", "-h", "--help"
				' Display usage
				Call Display_Usage("")
			Case "/e", "-e", "--export"
				' Don't export playlists from iTunes
				bExportFromiTunes = False
			Case "/d", "-d", "--delete"
				' Don't delete playlists from Plex
				bDeletePlexPlaylists = False
			Case "/u", "-u", "--upload"
				' Don't upload playlists to Plex
				bUploadToPlex = False
			Case "/s", "-s", "--smart"
				' Ignore smart playlists in iTunes
				bIgnoreSmartPlaylists = True
			Case "/v", "-v", "--verbose"
				' Verbose - show calls to Plex
				bVerbose = True
			Case Else
				Call Display_Usage("Unknown option (" & WScript.Arguments(iCount) & ")")
		End Select
	Next

End Sub

' Display_Usage
' Explain to the user how the script works. If sError contains a string
' then this is appended to the bottom.

Sub Display_Usage(sError)

	Dim sText
	Dim sName : sName = Wscript.ScriptName
	
	If Instr(sName, " ") <> 0 Then sName = """" & sName & """"
	
	sText = "Usage: cscript.exe " & sName & " [/E] [/S] [/D] [/U] [/V]" & VbCrLf
	sText = sText & VbCrLf
	sText = sText & "    No args     Start export from iTunes and upload to Plex." & VbCrLf
	sText = sText & "    /?          Display help." & VbCrLf
	sText = sText & "    /E          Don't export from iTunes." & VbCrLf
	sText = sText & "    /S          Don't export smart playlists from iTunes." & VbCrLf
	sText = sText & "    /D          Don't delete playlists from Plex." & VbCrLf
	sText = sText & "    /U          Don't upload to Plex." & VbCrLf
	sText = sText & "    /V          Verbose mode. Show commands executed."
	
	If sError <> "" Then
		sText = sText & VbCrLf & VbCrLf & "Error: " & sError
	End If

	WScript.Echo sText
	WScript.Quit

End Sub

' Log(sMessage)
' Display a message from the program on the screen. In the future this can be updated
' to write the logs out to a file and support verbosity settings.

Sub Log(sMessage)
	
	Dim sEntry : sEntry = "[" & FormatDateTime(Now(), vbLongTime) & "] " & sMessage

	' If the Plex token is in the string, then replace it with stars of equal length. This
	' is to avoid accidentally leaking the Plex token
	If Len(TOKEN) > 0 And Instr(sEntry, TOKEN) > 0 Then sEntry = Replace(Space(Len(TOKEN)), " ", "*")
	
	WScript.Echo sEntry
		
End Sub

' Test_Plex
' Test to see if Plex can be found and the token works. Returns True if it works
' and False if it doesn't.

Function Test_Plex

	Test_Plex = False

	' Token is empty

	If TOKEN = "" Then
		Log("Plex connection test failed: token is missing")
		Exit Function
	End If

	Dim sOutput : sOutput = Execute_Command("curl -sS """ & SERVER & "?X-Plex-Token=" & TOKEN & """", False)

	' Server and token look good

	If Instr(sOutput, "MediaContainer") > 0 Then
		Test_Plex = True
		Log("Plex connection test successful")
		Exit Function
	End If

	' Either the server or the token is wrong

	If Instr(sOutput, "401 Unauthorized") > 0 Then
		Log("Plex connection test failed: invalid token")
	Else
		Log("Plex connection test failed: can't connect to " & SERVER)
	End If
	
End Function

' Get_Playlist_Name
' Given a path and filename to a m3u, look for the #PLAYLIST: tag within
' the first 5 lines and return it.

Function Get_Playlist_Name(sFilename)

	Dim oStream : Set oStream = CreateObject("ADODB.Stream")
	
	With oStream
		.Charset = "utf-8"
		.Type = 2 ' adTypeText
		.LineSeparator = 10 ' adLF so we can handle Linux too
		.Open
		.LoadFromFile sFilename
	End With

	On Error Resume Next
		
	' Look at first 5 lines of a m3u file for #PLAYLIST:
		
	Dim iPos, sLine
	
	For iPos = 0 To 5
		' Read the line, remove any CRs and trim whitespace
		sLine = Trim(Replace(oStream.ReadText(-2), VbCr, ""))
		' Check if we have the playlist line
		If Len(sLine) > 10 And Left(LCase(sLine), 10) = "#playlist:" Then
			Get_Playlist_Name = Mid(sLine, 11, Len(sLine))
			If Get_Playlist_Name <> "" Then Exit For
		End If
	Next
		
	On Error Goto 0
	oStream.Close
	Set oStream = Nothing
	
	' If there isn't one in the file then take it from the filename
	If Get_Playlist_Name = "" Then Get_Playlist_Name = fso.GetBaseName(sFilename)

End Function