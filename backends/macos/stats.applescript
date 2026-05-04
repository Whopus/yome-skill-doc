-- doc stats — words, paragraphs, characters, pages
set TAB_CHAR to (ASCII character 9)
tell application "Microsoft Word"
    tell active document
        set wc to compute statistics statistic statistic words
        set pc to count of paragraphs
        set cc to compute statistics statistic statistic characters
        set pgc to 0
        try
            set pgc to compute statistics statistic statistic pages
        end try
        return (wc as string) & TAB_CHAR & (pc as string) & TAB_CHAR & (cc as string) & TAB_CHAR & (pgc as string)
    end tell
end tell
