-- doc comments — list comments (TSV: index, author, text)
set TAB_CHAR to (ASCII character 9)
set LF_CHAR to (ASCII character 10)
set rows to {"index" & TAB_CHAR & "author" & TAB_CHAR & "text"}

set authors to {}
set texts to {}
tell application "Microsoft Word"
    tell active document
        set n to count of comments
        repeat with i from 1 to n
            set c to comment i
            set end of authors to (author of c) as string
            set txt to ""
            try
                set txt to (comment text of c) as string
            end try
            -- Strip newlines/tabs from text
            set AppleScript's text item delimiters to TAB_CHAR
            set tmp to text items of txt
            set AppleScript's text item delimiters to " "
            set txt to tmp as string
            set AppleScript's text item delimiters to LF_CHAR
            set tmp to text items of txt
            set AppleScript's text item delimiters to " "
            set txt to tmp as string
            set AppleScript's text item delimiters to ""
            set end of texts to txt
        end repeat
    end tell
end tell
repeat with i from 1 to count of authors
    set end of rows to (i as string) & TAB_CHAR & (item i of authors) & TAB_CHAR & (item i of texts)
end repeat
set AppleScript's text item delimiters to LF_CHAR
set out to rows as string
set AppleScript's text item delimiters to ""
return out
