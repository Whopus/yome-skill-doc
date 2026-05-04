-- doc close [--save=true|false]
set saveSpec to {{save|json}}
set savingClause to "saving yes"
if saveSpec is "false" then set savingClause to "saving no"

tell application "Microsoft Word"
    if (count of documents) is 0 then return "no document"
    if savingClause is "saving yes" then
        try
            set fp to full name of active document
        on error
            return "ERROR: active document has no file path; use --save=false to discard"
        end try
    end if
    close active document saving (savingClause is "saving yes")
end tell
return "closed"
