-- doc image.add <path> [--index=<para_idx>]
set imgPath to {{path|json}}
set idxStr to {{index|json}}
set imgFile to POSIX file imgPath as string

tell application "Microsoft Word"
    tell active document
        if idxStr is "" then
            set anchor to end of text object
        else
            set anchor to text object of paragraph ((idxStr as integer))
        end if
        make new inline picture at anchor with properties {file name:imgFile}
    end tell
end tell
return "added"
