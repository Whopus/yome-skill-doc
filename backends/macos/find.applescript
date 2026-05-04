-- doc find <what> [--matchCase] [--wholeWord]
set whatStr to {{what|json}}
set caseFlag to {{matchCase|bool}}
set wordFlag to {{wholeWord|bool}}

tell application "Microsoft Word"
    tell active document
        set fo to find object of text object
        try
            clear formatting fo
        end try
        set found to false
        try
            set found to (execute find fo find text whatStr match case caseFlag match whole word wordFlag)
        end try
        if found is false then return "not found"
        set sel to text object of selection
        set s to start of content of sel
        set e to end of content of sel
        return "found at offset " & (s as string) & ".." & (e as string)
    end tell
end tell
