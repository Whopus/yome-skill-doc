-- doc replace.all <what> --with=<text> [--matchCase] [--wholeWord]
set whatStr to {{what|json}}
set withStr to {{with|json}}
set caseFlag to {{matchCase|bool}}
set wordFlag to {{wholeWord|bool}}

set countReplaced to 0
tell application "Microsoft Word"
    tell active document
        set fo to find object of text object
        try
            clear formatting fo
        end try
        repeat
            set didReplace to false
            try
                set didReplace to (execute find fo find text whatStr replace with withStr replace replace one match case caseFlag match whole word wordFlag)
            end try
            if didReplace then
                set countReplaced to countReplaced + 1
            else
                exit repeat
            end if
            if countReplaced > 10000 then exit repeat
        end repeat
    end tell
end tell
return "replaced " & (countReplaced as string)
