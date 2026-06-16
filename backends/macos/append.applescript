-- doc append --text=<content>
set theText to {{text|json}}
tell application "Microsoft Word"
    tell document 1
        if (count of paragraphs) is 1 then
            set firstText to content of text object of paragraph 1
            if firstText ends with return then
                if (length of firstText) > 1 then
                    set firstText to text 1 thru -2 of firstText
                else
                    set firstText to ""
                end if
            end if
            if firstText is "" then
                set content of text object of paragraph 1 to theText & return
                return "appended"
            end if
        end if
        set lastIdx to count of paragraphs
        set lastText to content of text object of paragraph lastIdx
        if lastText ends with return then
            if (length of lastText) > 1 then
                set lastText to text 1 thru -2 of lastText
            else
                set lastText to ""
            end if
        end if
        if lastText is "" then
            set content of text object of paragraph lastIdx to theText & return
        else
            set content of text object to (content of text object) & theText & return
        end if
    end tell
end tell
return "appended"
