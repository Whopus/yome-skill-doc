-- doc page [--orientation=portrait|landscape] [--paper=letter|a4|legal|...] [--margin=N (points)]
set orientStr to {{orientation|json}}
set paperStr to {{paper|json}}
set marginStr to {{margin|json}}

tell application "Microsoft Word"
    tell active document
        set ps to page setup
        if orientStr is "landscape" then
            set orientation of ps to page orientation landscape
        else if orientStr is "portrait" then
            set orientation of ps to page orientation portrait
        end if
        if paperStr is "letter" then set paper size of ps to paper size US letter
        if paperStr is "a4"     then set paper size of ps to paper size A4
        if paperStr is "legal"  then set paper size of ps to paper size US legal
        if paperStr is "a3"     then set paper size of ps to paper size A3
        if paperStr is "a5"     then set paper size of ps to paper size A5
        if marginStr is not "" then
            set m to (marginStr as real)
            try
                set top margin of ps to m
                set bottom margin of ps to m
                set left margin of ps to m
                set right margin of ps to m
            end try
        end if
    end tell
end tell
return "set"
