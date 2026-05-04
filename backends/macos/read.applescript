-- doc read <index> — full text + format JSON
set idx to ({{index|json}}) as integer
set TAB_CHAR to (ASCII character 9)

tell application "Microsoft Word"
    tell active document
        set p to paragraph idx
        set paraText to content of text object of p
        if paraText ends with return then
            set paraText to text 1 thru -2 of paraText
        end if
        set fo to font object of text object of p
        set isBold to (bold of fo) as string
        set isItalic to (italic of fo) as string
        set fSize to (font size of fo) as string
        set fName to name of fo
        set fColor to (color index of fo) as string
        return paraText & TAB_CHAR & isBold & TAB_CHAR & isItalic & TAB_CHAR & fSize & TAB_CHAR & fName & TAB_CHAR & fColor
    end tell
end tell
