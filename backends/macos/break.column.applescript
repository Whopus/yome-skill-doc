-- doc break.column [--index=<para_idx>]
set idxStr to {{index|json}}
tell application "Microsoft Word"
    tell active document
        if idxStr is "" then
            insert break at end of text object break type column break
        else
            insert break at text object of paragraph ((idxStr as integer)) break type column break
        end if
    end tell
end tell
return "inserted"
