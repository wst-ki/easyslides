-- Export a PPTX file to PDF via Microsoft PowerPoint on macOS.
-- Expects paths inside PowerPoint's sandbox-accessible Group Container so no
-- interactive "Grant File Access" dialog blocks unattended runs.
-- Intentionally avoids `activate` and never quits PowerPoint if it was already
-- open, minimizing disruption to any foreground work.

on run argv
    if (count of argv) < 2 then
        error "Usage: osascript powerpoint_export.applescript <input_pptx_posix> <output_pdf_posix>"
    end if

    set inputPosix to item 1 of argv
    set outputPosix to item 2 of argv

    tell application "System Events"
        set wasRunning to (name of processes) contains "Microsoft PowerPoint"
    end tell

    set inputFile to POSIX file inputPosix
    set openedPres to missing value

    tell application "Microsoft PowerPoint"
        try
            set openedPres to open inputFile read only yes with window no
        on error
            set openedPres to open inputFile read only yes
        end try

        if openedPres is missing value then
            set openedPres to active presentation
        end if

        save openedPres in outputPosix as save as PDF file
        close openedPres saving no

        if not wasRunning then
            quit
        end if
    end tell
end run
