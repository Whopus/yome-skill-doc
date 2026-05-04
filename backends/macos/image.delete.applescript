-- doc image.delete <index>
set idx to ({{index|json}}) as integer
tell application "Microsoft Word"
    tell active document
        delete inline picture idx
    end tell
end tell
return "deleted"
