-- doc bookmarks — list bookmarks (TSV: name, start)
set TAB_CHAR to (ASCII character 9)
set LF_CHAR to (ASCII character 10)
set rows to {"name" & TAB_CHAR & "start"}

set names to {}
set starts to {}
tell application "Microsoft Word"
    tell active document
        set bms to every bookmark
        repeat with b in bms
            set end of names to (name of b) as string
            try
                set end of starts to (start of content of text object of b) as string
            on error
                set end of starts to "0"
            end try
        end repeat
    end tell
end tell
repeat with i from 1 to count of names
    set end of rows to (item i of names) & TAB_CHAR & (item i of starts)
end repeat
set AppleScript's text item delimiters to LF_CHAR
set out to rows as string
set AppleScript's text item delimiters to ""
return out
