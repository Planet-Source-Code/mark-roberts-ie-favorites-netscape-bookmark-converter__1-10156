<div align="center">

## IE Favorites \<\-\> Netscape bookmark converter


</div>

### Description

Converts IE Favorites to Netscape Bookmarks and vice-versa in just 9 lines of code
 
### More Info
 
Just the Netscape bookmark.htm file path

To use this code, make sure you have an Internet Control component in your project. Then its simple - call the functions with the path of your netscape bookmark file.

The export operation will not merge with the Netscape file - it will overwrite it.


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Mark Roberts](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/mark-roberts.md)
**Level**          |Advanced
**User Rating**    |3.7 (11 globes from 3 users)
**Compatibility**  |VB 4\.0 \(16\-bit\), VB 4\.0 \(32\-bit\), VB 5\.0, VB 6\.0
**Category**       |[Internet/ HTML](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/internet-html__1-34.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/mark-roberts-ie-favorites-netscape-bookmark-converter__1-10156/archive/master.zip)





### Source Code

```
Dim ShellUIHelper1 As ShellUIHelper
Sub ImportFavorites(NetscapePath As String)
 Set ShellUIHelper1 = New ShellUIHelper
 ShellUIHelper1.ImportExportFavorites True, NetscapePath
End Sub
Sub ExportFavorites(NetscapePath As String)
 Set ShellUIHelper1 = New ShellUIHelper
 ShellUIHelper1.ImportExportFavorites False, NetscapePath
End Sub
```

