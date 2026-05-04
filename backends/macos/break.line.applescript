-- doc break.line [--index=<para_idx>]
set idxStr to {{index|json}}
tell application "Microsoft Word"
    tell active document
        if idxStr is "" then
            insert break at end of text object break type line break
        else
            insert break at text object of paragraph ((idxStr as integer)) break type line break
        end if
    end tell
end tell
return "inserted"
