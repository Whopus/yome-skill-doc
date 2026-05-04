-- doc break.page [--index=<para_idx>]   (default: end of doc)
set idxStr to {{index|json}}
tell application "Microsoft Word"
    tell active document
        if idxStr is "" then
            set anchor to end of text object
            insert break at anchor break type page break
        else
            set anchor to start of content of text object of paragraph ((idxStr as integer))
            insert break at text object of paragraph ((idxStr as integer)) break type page break
        end if
    end tell
end tell
return "inserted"
