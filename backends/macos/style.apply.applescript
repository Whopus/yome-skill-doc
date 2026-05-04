-- doc style.apply <index> --style=<style_name>
set idx to ({{index|json}}) as integer
set styleName to {{style|json}}

tell application "Microsoft Word"
    tell active document
        set st to style styleName of active document
        set style of paragraph idx to st
    end tell
end tell
return "applied"
