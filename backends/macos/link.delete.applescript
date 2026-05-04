-- doc link.delete <index>
set idx to ({{index|json}}) as integer
tell application "Microsoft Word"
    tell active document
        set para to text object of paragraph idx
        set ps to start of content of para
        set pe to end of content of para
        try
            delete (every hyperlink whose start of content of text object >= ps and end of content of text object <= pe)
        end try
    end tell
end tell
return "deleted"
