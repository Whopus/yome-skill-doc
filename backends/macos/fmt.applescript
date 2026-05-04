-- doc fmt <index> [--bold] [--italic] [--size] [--color] [--font] [--align]
set idx to ({{index|json}}) as integer
set boldFlag to {{bold|bool}}
set italicFlag to {{italic|bool}}
set sizeStr to {{size|json}}
set colorStr to {{color|json}}
set fontStr to {{font|json}}
set alignStr to {{align|json}}

tell application "Microsoft Word"
    tell active document
        set p to paragraph idx
        set fo to font object of text object of p

        if boldFlag then
            try
                set bold of fo to true
            end try
        end if
        if italicFlag then
            try
                set italic of fo to true
            end try
        end if
        if sizeStr is not "" then
            try
                set font size of fo to (sizeStr as real)
            end try
        end if
        if fontStr is not "" then
            try
                set name of fo to fontStr
            end try
        end if
        if colorStr is not "" then
            try
                set color of fo to {{color|rgb}}
            end try
        end if
        if alignStr is not "" then
            try
                set paragraph alignment of paragraph format of text object of p to {{align|align}}
            end try
        end if
    end tell
end tell
return "updated"
