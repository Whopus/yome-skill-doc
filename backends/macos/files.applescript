-- doc files — list open documents (TSV: name, paragraphs)
set TAB_CHAR to (ASCII character 9)
set LF_CHAR to (ASCII character 10)

set rows to {"name" & TAB_CHAR & "paragraphs"}
tell application "Microsoft Word"
    set n to count of documents
    repeat with i from 1 to n
        set d to document i
        set dn to name of d
        set pc to count of paragraphs of d
        set end of rows to dn & TAB_CHAR & (pc as string)
    end repeat
end tell
set AppleScript's text item delimiters to LF_CHAR
set out to rows as string
set AppleScript's text item delimiters to ""
return out
