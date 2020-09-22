<div align="center">

## ASP File Finder


</div>

### Description

Did you ever want to search for files using your web browser instead of the MS Find Files program? This ASP file searches your hard drive (or web server) for files containing a given string. You can specify a string to search for and the directory to search in (or leave the default c:\ directory).
 
### More Info
 
SearchText = The file containing the text you are searching for.

Directory = The Directory (and it's subdirectories) to search (default is c:\)

I realize this app is much slower than the MS Find Files program, but it is web enabled and will not return files that do not contain the search text like the MS program does.

A list of files on your hard drive (or web server) matching the string you searched for.

On slow machines - This script may time out if you search the full directory. Note: On my 1GHz machine it can take up to 2 minutes to search the entire 34GB hard drive (I set the script timeout to 5 minutes).


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Blowno](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/blowno.md)
**Level**          |Intermediate
**User Rating**    |4.5 (18 globes from 4 users)
**Compatibility**  |ASP \(Active Server Pages\), HTML, VbScript \(browser/client side\)

**Category**       |[Files](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/files__4-2.md)
**World**          |[ASP / VbScript](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/asp-vbscript.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/blowno-asp-file-finder__4-6346/archive/master.zip)

### API Declarations

NOTE: The code that traverses the folders was lifted from another app from the PSC site so I am not posting this for purposes of obtaining votes, but just for informational purposes.


### Source Code

```
<%@LANGUAGE="VBSCRIPT"%>
<%
  Response.AddHeader "Pragma", "No-Cache"	'try not to cache page
  Response.CacheControl = "Private"		'try not to cache page
	server.scripttimeout = 300		'script will time out after 5 minutes
%>
<html>
<head><title>Find Files Using ASP</title>
<style>
 body {font:10pt Arial;background-color:papayawhip;color:antiquewhite;font-weight:bold;margin-top:0px;margin-left:0px;margin-right:0px}
 A:link {color:black;text-decoration:none}
 A:hover {color:red;text-decoration:underline}
 A:visited {color:black;text-decoration:none}
 td {color:black;border-bottom:1pt solid black;font:9pt Arial}
 th {color:black;border-bottom:1pt solid black;font:9pt Arial;font-weight:bold}
</style>
</head>
<body>
<div style="background-color:tan">
<center>
Find Files<br>
<%
dim filecounter, searchtext, directory	'dimention variables
dim fcount, fsize
filecounter = 0				'initialize filecounter to zero
searchtext = Trim(request("SearchText"))	'get the querystring SearchText
directory = Trim(request("Directory"))		'get the querystring Directory
if directory = "" then directory = "c:\"	'if no directory the set to c:\
						'Write the Search Form to the page
response.write "<form action='FindFiles.asp' method=get>Search For:" & _
	" <input type=text name=SearchText size=20 value=" & Chr(34) & searchtext & _
	Chr(34) & "> Change Directory: <input type=text name=Directory size=20 value=" & _
	Chr(34) & directory & Chr(34) & "> <input style='background-color:blanchedalmond;" & _
	"color:chocolate' type=submit value='Start Search'></form><br></div>"
if searchtext <> "" Then	'if there is a file to search for then search
response.write "<table border=0 width='100%'>"
response.write "<tr><th width='60%'>File Name</th><th width='10%'>File Size</th><th width='30%'>Date Modified</th></tr>"
		'create the recordset object to store
		'the filepath, filename, filesize and last modified date
     set rs = createobject("adodb.recordset")
     rs.fields.append "FilePath",200,255
	 rs.fields.append "FileName",200,255
     rs.fields.append "FileSize",200,255
	 rs.fields.append "FileDate",7,255
     rs.open
	Recurse directory	'call the subroutine to traverse the directories
  	Sub Recurse(Path)
			'create the file system object
  		Dim fso, Root, Files, Folders, File, i, FoldersArray(1000)
  		Set fso = Server.CreateObject("Scripting.FileSystemObject")
  		Set Root = fso.getfolder(Path)
  		Set Files = Root.Files
  		Set Folders = Root.SubFolders
		fcount = 0			'zero out the file count variable
			'traverse through the subdirectories in the current directory
  		For Each Folder In Folders
  			FoldersArray(i) = Folder.Path
  			i = i + 1
  		Next
			'traverse through the files in the current folder or subfolder
  		For Each File In Files
				'check if the search string is found
			num = InStr(UCase(File.Name), UCase(searchtext))
				'if it is then update the recordset and sort it
			if num <> 0 then
			filecounter = filecounter + 1
			rs.addnew
		    rs.fields("FilePath") = File.Path
			rs.fields("FileName") = File.Name
			rs.fields("FileSize") = File.Size
			rs.fields("FileDate") = File.DateLastModified
			rs.update
       		rs.Sort = "FileName ASC"
			end if
  		Next
			'recurse through the current directory until
			'all subfolders have been traversed
  		For i = 0 To UBound(FoldersArray)
  			If FoldersArray(i) <> "" Then
  				Recurse FoldersArray(i)
  			Else
  				Exit For
  			End If
  		Next
  	End Sub
		'if files were found then write them to the document
	If filecounter <> 0 then
			filecounter = 0
		do while not rs.eof
			filecounter = filecounter + 1
			response.write "<tr><td width='50%' valign=top><a href=""" & rs.fields("FilePath") & """>" & rs.fields("FileName") & "</td><td width='10%' align=right valign=top>"
					'get the file size so we can
					'assign the proper Bytes, KB or MB value
				fsize = CLng(rs.fields("FileSize"))
				'if less than 1 kilobyte then it's Bytes
			if fsize >= 0 And fsize <= 999 then
				fnumber = FormatNumber(fsize,0) & " Bytes"
			end if
				'if 1 KB but less then 1 MB then assign KB
			if fsize >= 1000 And fsize <= 999999 then
				fnumber = FormatNumber((fsize / 1000),2) & " KB"
			end if
				'if 1 MB or more then assign MB
			if fsize >= 1000000 then
				fnumber = FormatNumber((fsize / 1000000),2) & " MB"
			end if
				'write each file and corresponding info to the document
			response.write fnumber & "</td><td width='30%' align='center'>" & rs.fields("FileDate") & "</td></tr>"
			rs.movenext
		loop
		response.write "</table>"	'end the table
	else
			'no files were found
	end if
end if
%>
</body>
</html>
```

