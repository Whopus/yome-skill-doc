-- doc bookmark.delete <name>
set bmName to {{name|json}}
tell application "Microsoft Word"
    tell active document
        delete (every bookmark whose name is bmName)
    end tell
end tell
return "deleted"
