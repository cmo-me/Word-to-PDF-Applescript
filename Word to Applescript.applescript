-- WATCHOUT - DO NOT USE ON PUBLIC FOLDERS AS THE SCRIPT CAN USE ALREADY OPEN DOCUMENTS VS THE INTENDED DOCUMENT -- IT HAS A BUG

-- this script adapted from post at https://discussions.apple.com/thread/3050596?start=0&tstart=0 solution by spazek
-- details on the use of this script for Hazel found a http://scrubbs.me
--  Updated with delay to solve issue as described by Corwin Carr   at Mac Power Users Show 186 notes http://5by5.tv/mpu/186  
-- the First two commands avoid an error where Word would be unable to get the file path. The script would ultimately run appropriately, but only after throwing the error. 

tell application "Microsoft Word" to activate
delay 1

tell application "Microsoft Word" to set theOldDefaultPath to get default file path file path type documents path -- looks like we change the default path to where the document is and then set it back when we're done
try
    tell application "Finder"
        set theFilePath to container of theFile as text
        
        set ext to name extension of theFile
        
        set theName to name of theFile
        copy length of theName to l
        copy length of ext to exl
        
        set n to l - exl - 1
        copy characters 1 through n of theName as string to theFilename
        
        set theFilename to theFilename & ".pdf"
        
        tell application "Microsoft Word"
            set default file path file path type documents path path theFilePath
            open theFile
            set theActiveDoc to the active document
            save as theActiveDoc file format format PDF file name theFilename
            close theActiveDoc
        end tell
        
    end tell
    
end try
try
    
    tell application "Microsoft Word" to set default file path file path type documents path path theOldDefaultPath
    
end try
