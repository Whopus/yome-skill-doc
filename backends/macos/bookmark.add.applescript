-- doc bookmark.add <index> --name=<bookmark_name>
set idx to ({{index|json}}) as integer
set bmName to {{name|json}}
tell application "Microsoft Word"
    tell active document
        make new bookmark at text object of paragraph idx with properties {name:bmName}
    end tell
end tell
return "added"
