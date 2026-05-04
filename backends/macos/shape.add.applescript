-- doc shape.add [--kind=rectangle|oval|...] [--left=N --top=N --width=N --height=N] [--text=<str>]
set leftPos to ({{left|json}}) as real
set topPos to ({{top|json}}) as real
set widthVal to ({{width|json}}) as real
set heightVal to ({{height|json}}) as real
set bodyText to {{text|json}}

tell application "Microsoft Word"
    tell active document
        set sh to make new shape at end with properties {auto shape type:{{kind|autoshape}}, left position:leftPos, top:topPos, width:widthVal, height:heightVal}
        if bodyText is not "" then
            try
                set content of text object of text frame of sh to bodyText
            end try
        end if
        return name of sh
    end tell
end tell
