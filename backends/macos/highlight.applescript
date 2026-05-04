-- doc highlight <index> [--color=yellow|green|pink|blue|red|gray]
set idx to ({{index|json}}) as integer
set colorName to {{color|json}}
if colorName is "" then set colorName to "yellow"

set hi to yellow color
if colorName is "green" then set hi to bright green color
if colorName is "pink" then set hi to pink color
if colorName is "blue" then set hi to turquoise color
if colorName is "red" then set hi to red color
if colorName is "gray" then set hi to gray 25 color
if colorName is "grey" then set hi to gray 25 color
if colorName is "none" then set hi to no highlight color

tell application "Microsoft Word"
    tell active document
        try
            set highlight color index of text object of paragraph idx to hi
        end try
    end tell
end tell
return "highlighted"
