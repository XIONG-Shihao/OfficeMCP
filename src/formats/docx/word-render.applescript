on closeOwned(inputPath, outputPath, applicationPath)
  set inputHFS to (POSIX file inputPath) as text
  set outputHFS to (POSIX file outputPath) as text
  using terms from application "Microsoft Word"
  tell application applicationPath
    repeat with docOrdinal from (count of documents) to 1 by -1
      set ownedDocument to document (docOrdinal as integer)
      set actualPath to (full name of ownedDocument) as text
      if actualPath is inputPath or actualPath is outputPath or actualPath is inputHFS or actualPath is outputHFS then
        close ownedDocument saving no
      end if
    end repeat
  end tell
  end using terms from
end closeOwned

on run argv
  set inputPath to item 1 of argv
  set outputPath to item 2 of argv
  set phase to item 3 of argv
  set applicationPath to item 4 of argv
  if phase is "open" then
    my closeOwned(inputPath, outputPath, applicationPath)
    using terms from application "Microsoft Word"
  tell application applicationPath
      open file name inputPath read only true add to recent files false
      return version
    end tell
  end using terms from
  end if
  if phase is "close" then
    my closeOwned(inputPath, outputPath, applicationPath)
    return "closed"
  end if
  set inputHFS to (POSIX file inputPath) as text
  using terms from application "Microsoft Word"
  tell application applicationPath
    set openedDocument to missing value
      repeat with docOrdinal from 1 to (count of documents)
        set checkedDocument to document (docOrdinal as integer)
        set checkedPath to (full name of checkedDocument) as text
        if checkedPath is inputPath or checkedPath is inputHFS then
          set openedDocument to checkedDocument
          exit repeat
        end if
      end repeat
    if openedDocument is missing value then error "OfficeMCP owned input is not ready." number -2701
    try
      save as openedDocument file name outputPath file format format PDF add to recent files false
      my closeOwned(inputPath, outputPath, applicationPath)
      return version
    on error errorMessage number errorNumber
      try
        my closeOwned(inputPath, outputPath, applicationPath)
      end try
      error errorMessage number errorNumber
    end try
  end tell
  end using terms from
end run
