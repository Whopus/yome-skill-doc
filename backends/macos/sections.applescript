-- doc sections — list sections (TSV: index, paragraphs)
set TAB_CHAR to (ASCII character 9)
set LF_CHAR to (ASCII character 10)
set rows to {"index" & TAB_CHAR & "paragraphs"}

tell application "Microsoft Word"
    tell active document
        set n to count of sections
        repeat with i from 1 to n
            set s to section i
            set pc to count of paragraphs of text object of s
            set end of rows to (i as string) & TAB_CHAR & (pc as string)
        end repeat
    end tell
end tell
set AppleScript's text item delimiters to LF_CHAR
set out to rows as string
set AppleScript's text item delimiters to ""
return out
