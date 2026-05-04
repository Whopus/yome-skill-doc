-- doc break.section [--index=<para_idx>] [--kind=continuous|next-page|even-page|odd-page]
set idxStr to {{index|json}}
set kindStr to {{kind|json}}
if kindStr is "" then set kindStr to "next-page"

set bt to section break next page
if kindStr is "continuous" then set bt to section break continuous
if kindStr is "even-page"  then set bt to section break even page
if kindStr is "odd-page"   then set bt to section break odd page

tell application "Microsoft Word"
    tell active document
        if idxStr is "" then
            insert break at end of text object break type bt
        else
            insert break at text object of paragraph ((idxStr as integer)) break type bt
        end if
    end tell
end tell
return "inserted"
