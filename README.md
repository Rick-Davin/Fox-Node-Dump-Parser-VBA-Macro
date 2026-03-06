# Fox-Node-Dump-Parser-VBA-Macro
Legacy XLS macro to transform a Fox Node Dump text file to an Excel Workbook.

Users of Invensys Foxboro IA Distributed Control Systems (DCS) may produce a node dump, which is a text file.  If this text file is copied to a Windows PC running Excel, the then macro Fox2Excel will read the node dump and transform the data to spreadsheets.  Each block on information generally contains a NAME and TYPE attribute, as well as many other attributes associated with the TYPE.  Each attribute is written to a single line as a Name-Value pair, though it literally would be "NAME = value".  That is to say the attribute name is followed by a blank, an equal sign, another blank, and lastly the value.  A block begins with a "NAME = " line, though that phrase may be preceded by several blanks.  A block ends on a line, when trimmed, simply equals "END".

Each sheet is named the same as the TYPE.

The Fox attributes are typically uppercase, such as NAME, TYPE, LOOPID, etc.  The values for the TYPE, which is the name of the sheets, are also uppercased and should be less than 25 characters.  The macro generates some meta data columns, which will begin with an underscore and be lowercased.  Likewise, a meta sheet would also have lowercased names prefaced with an underscore.  There is some specific logic when looping through the node data that sheets beginning with underscore will be skipped.

