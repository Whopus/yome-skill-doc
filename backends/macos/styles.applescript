-- doc styles — list document styles (TSV: name, type)
set TAB_CHAR to (ASCII character 9)
set LF_CHAR to (ASCII character 10)
set rows to {"name" & TAB_CHAR & "type"}

set names to {}
set types to {}
tell application "Microsoft Word"
    tell active document
        set sts to every style
        repeat with st in sts
            set end of names to (name local of st) as string
            try
                set end of types to (type of st) as string
            on error
                set end of types to ""
            end try
        end repeat
    end tell
end tell
repeat with i from 1 to count of names
    set end of rows to (item i of names) & TAB_CHAR & (item i of types)
end repeat
set AppleScript's text item delimiters to LF_CHAR
set out to rows as string
set AppleScript's text item delimiters to ""
return out
