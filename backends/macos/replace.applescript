-- doc replace <what> --with=<text> [--matchCase] [--wholeWord]  (first match only)
set whatStr to {{what|json}}
set withStr to {{with|json}}
set caseFlag to {{matchCase|bool}}
set wordFlag to {{wholeWord|bool}}

tell application "Microsoft Word"
    tell active document
        set fo to find object of text object
        try
            clear formatting fo
        end try
        set didReplace to false
        try
            set didReplace to (execute find fo find text whatStr replace with withStr replace replace one match case caseFlag match whole word wordFlag)
        end try
        if didReplace then
            return "replaced 1"
        else
            return "no match"
        end if
    end tell
end tell
