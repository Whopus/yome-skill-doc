-- doc get — list paragraphs (TSV: index, preview)
set TAB_CHAR to (ASCII character 9)
set LF_CHAR to (ASCII character 10)
set out to "index" & TAB_CHAR & "text"

tell application "Microsoft Word"
    tell active document
        set pc to count of paragraphs
        repeat with i from 1 to pc
            set paraText to content of text object of paragraph i
            if paraText ends with return then
                if (length of paraText) > 1 then
                    set paraText to text 1 thru -2 of paraText
                else
                    set paraText to ""
                end if
            end if
            if (length of paraText) > 80 then
                set preview to (text 1 thru 80 of paraText) & "..."
            else
                set preview to paraText
            end if
            -- Strip embedded tabs/newlines from preview to keep TSV intact
            set AppleScript's text item delimiters to TAB_CHAR
            set tmp to text items of preview
            set AppleScript's text item delimiters to " "
            set preview to tmp as string
            set AppleScript's text item delimiters to LF_CHAR
            set tmp to text items of preview
            set AppleScript's text item delimiters to " "
            set preview to tmp as string
            set AppleScript's text item delimiters to ""
            set out to out & LF_CHAR & (i as string) & TAB_CHAR & preview
        end repeat
    end tell
end tell

return out
