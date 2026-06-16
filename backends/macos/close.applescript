-- doc close [--save=true|false]
set saveSpec to {{save|json}}
set savingClause to "saving yes"
if saveSpec is "false" then set savingClause to "saving no"

tell application "Microsoft Word"
    if (count of documents) is 0 then return "no document"
    set beforeCount to count of documents
    if savingClause is "saving yes" then
        try
            set fp to full name of document 1
        on error
            return "ERROR: active document has no file path; use --save=false to discard"
        end try
        save document 1
    else
        set saved of document 1 to true
    end if
    activate
end tell

delay 0.2
tell application "System Events"
    key code 13 using command down
end tell
delay 0.5

tell application "Microsoft Word"
    if (count of documents) < beforeCount then return "closed"
end tell
return "ERROR: close failed; document is still open"
