-- doc append --text=<content>
set theText to {{text|json}}
tell application "Microsoft Word"
    tell active document
        set newPara to make new paragraph at end of text object of active document
        set content of text object of newPara to theText & return
    end tell
end tell
return "appended"
