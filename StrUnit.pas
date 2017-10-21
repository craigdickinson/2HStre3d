{'*******************************************************************************************
'  Name    : StrUnit.pas
'
'  Author  : Michael Campbell
'  Notice  : Copyright (c) 2009 2H Offshore Engineering Ltd 
'          : All Rights Reserved 
 '
'  Notes   : The main display screen code
'
' 
' CVS VERSION INFO
'
'      Last Checked in by: $Author: dickinsc $ 
'      Last Checked in:    $Date: 2014-03-12 09:31:50 $
'              Filename:   $RCSfile: StrUnit.pas,v $ 
'              Revision:   $Revision: 1.3 $
'              Tag:        $Name: not supported by cvs2svn $
'******************************************************************************************* }{ ******************************************************************* }
// *******************************************************************
//    Unit: String.Pas
//    Name: Christopher Bridge
//    Date: 20/7/99
//    Description: Procedures for String handling
// *******************************************************************

//	Revision History:
//	Version 	Name		     Date		     Description
//	1.0		C.Bridge	     20-07-99	     First implementation
//  1.1   M.Campbell     23-09-02      FirstWord updated for Delphi
//	1.2		C.Bridge	     12-01-06      FirstWord overloaded to except new string seperators, and MonthStr2Num function added.
//  1.3   T.Yang         16-11-12      Add "|" as word delimeter and add extra argument for function "FirstWord" to allow user use "-" in file name.
//  1.4   R.Harrison     13-02-13      FirstWord Update to allow comments to be included with "Delphi Brackets" as comments
//  1.4.1 R.Harrison     18-02-13      Case Sensitive Correction to allow to work with lower case input
//  1.4.2 R.Harrison     19-06-13      Find Line comment out...
//  1.4.3 C.Bowden       03/12/15      FirstWord range check error fixed 

unit StrUnit;

interface

  procedure FirstWord(var isFileName: boolean; var TextLine: String; var First: String); overload;
  procedure FirstWord(var isFileName: boolean; var TextLine: String; var First: String; const NewString: String); overload;
  procedure FirstWord(var TextLine: String; var First: String); overload;
  procedure FirstWord(var Textline: String; var First: String; const NewString: String); overload;
  procedure WordDot(TextLine: String; var IsDot: Integer);
  procedure Search(var OpenFile: Text; Const SearchString: String; var Result: String; NoReset : boolean = False;
  								 NotInComments : boolean = True; CaseSensitive : boolean = False);
  procedure SearchSU(var OpenFile: Text; Const SearchString: String; var Result: String);
  procedure SearchFirstWord(var OpenFile: Text; Const SearchString: String; var Result: String);
  procedure SearchKeyWord(var OpenFile: Text; Const SearchString: String; Const SearchChar: Char; var Result: String);
  procedure SearchBegin(var OpenFile: Text; Const SearchString: String; var Result: String;
												Characters: Integer; var Counter: Integer);
  procedure SearchNextLine(var OpenFile: Text; var Result: String);
  procedure SearchToNextNumber(var OpenFile: Text; var Result: String);
  procedure SearchToNextNonNumber(var OpenFile: Text; var Result: String);
  function UpStr(Const TextLine: String): String;
  function FindLine(var searchFile: TextFile; searchStr: String; fileReset: Boolean = True;
										notInComments: Boolean = True; caseSensitive: Boolean = False): String;
  function MonthStr2Num(const MonthLine: String): Integer;

implementation

uses
	SysUtils,
  StrUtils;

// *******************************************************************
// procedure FirstWord
// procedure to return first numonic word in string
// Note that comments in "Delphi bracket" will be ingored
//	Note: Numonic word is deleted from original string
// *******************************************************************
procedure FirstWord(var isFileName: Boolean; var TextLine: String; var First: String);
begin
	FirstWord(isFileName, TextLine, First, ' ');
end;

procedure FirstWord(var TextLine: String; var First: String);
var isFileName: Boolean;
begin
	isFileName := True;
  FirstWord(isFileName, TextLine, First, ' ');
end;

procedure FirstWord(var Textline: String; var First: String; const NewString: String);
var isFileName: Boolean;
begin
	isFileName := True;
  FirstWord(isFileName, Textline, First, NewString);
end;

procedure FirstWord(var isFileName: Boolean; var TextLine: String; var First: String; const NewString: String);
{
Procedure to return the first word in a string

isFileName: Boolean to determine whether or not to check for missing delimiter before a'-' sign
TextLine: String to search for first word
First: String to return first word
NewString: User defined delimiter - must be a string of length 1

The first word, any preceding delimiters and the first following delimiter are removed from TextLine
If a zero length string is passed then First is returned as a single space.

Delimiters:
  ' ' - space
  ',' - comma
  #09 - tab
  NewString - user defined

}
var
  Counter: Integer; // General Loop counting integer
  StringLength: Integer;
  TextStart, TextLength: Integer; // Start and length of 'First'
  CommentCounter, CommentCounter2, RemainingLength: Integer;
  StartofTextLine, EndofTextLine: String; {remaing section of TextLine }
  MinusCk, EndLoop, firstCharFound: Boolean;

begin

  if (Length (TextLine) > 0) and (TextLine <> ' ') then
  begin
    // Initialise Varibles etc
    firstCharFound := False;
    Counter := 0;
    StringLength := Length(TextLine);

	  // Loop until Start of first word is found
  	repeat
      Counter:= Counter + 1;
      if ((TextLine[Counter] <> ' ') and
          (TextLine[Counter] <> ',') and
          (TextLine <> #09 ) and
          (TextLine[Counter] <> NewString)) then
        firstCharFound := True;
		until (Counter=StringLength) or (firstCharFound=True);

    // Update the string cutting variables
    TextStart := Counter;

    // Setup Loop Variables
    EndLoop := False;
    MinusCk := False;

    // No words in string
    if firstCharFound=False then
    begin
      TextLine := '';
      First := ' ';
    end

    // Get first word if one exists
    else
    begin
		  // Find length of first string
		  repeat

        // remove the comment section from TextLine
				if TextLine[Counter] = '{' then
          repeat
            CommentCounter := 0;

            // Find the end of the comment
            repeat
              If CommentCounter+Counter < StringLength then
                Inc(CommentCounter,1);
            until (TextLine[CommentCounter+Counter] = '}') or (CommentCounter+Counter = StringLength);

            // Read to next word or comment
            CommentCounter2 := 0;                 
            if (CommentCounter+Counter<StringLength) then
              repeat
                inc(CommentCounter2);
              until (CommentCounter+Counter+CommentCounter2 = StringLength) or
                  ((TextLine[CommentCounter+Counter+CommentCounter2]<>' ') and
                   (TextLine[CommentCounter+Counter+CommentCounter2]<>',') and
                   (TextLine[CommentCounter+Counter+CommentCounter2]<>#09) and
                   (TextLine[CommentCounter+Counter+CommentCounter2]<>NewString));

            // Check if another word or comment was found
            if ((commentCounter2>0) and
                (TextLine[CommentCounter+Counter+CommentCounter2]<>' ') and
                (TextLine[CommentCounter+Counter+CommentCounter2]<>',') and
                (TextLine[CommentCounter+Counter+CommentCounter2]<>#09) and
                (TextLine[CommentCounter+Counter+CommentCounter2]<>NewString)) then
               CommentCounter2:=CommentCounter2-1;

            RemainingLength := Stringlength - (Counter+CommentCounter+commentCounter2);

            // Text after comment
            if RemainingLength > 0 then
              EndofTextLine := Copy(TextLine, Counter+CommentCounter+CommentCounter2+1, RemainingLength)
            else
              EndofTextLine := ' ';

            // Text before comment
            if Counter > 1 then
              StartofTextLine := Copy(TextLine, 1, Counter-1)
            else
              StartofTextLine := '';

            // String without comment
            TextLine := StartofTextLine + EndofTextLine;
            Stringlength := Length(TextLine);

          until (TextLine[Counter] <> '{');

        // Next character
        if counter<stringLength then
          Counter:= Counter + 1;

        if (Counter < StringLength) and (Counter > 1) then
        begin
          // Check for a minus sign that is too close to the previous number
          // and is not an exponent
          if (TextLine[Counter] = '-') and not isFileName then
          begin
            if (TextLine[Counter - 1] <> 'E') and
               (TextLine[Counter - 1] <> 'e') and
               (TextLine[Counter - 1] <> '-') then
            begin
              // Counter:= Counter - 1;
              EndLoop := True;
              MinusCk := True;
            end;
          end;
        end;

      // Stop at next delimiter
      until (Counter = StringLength) or
            (TextLine[Counter] = ' ') or
            (TextLine[Counter] = ',') or
            (TextLine[Counter] = #09) or
            (TextLine[Counter] = NewString) or
            (EndLoop = True);

      // Calculate the string length
      TextLength := Counter - TextStart;

      // Case that last character is not a delimiter or string is a single delimiter
      if ((Counter=StringLength) and
         (TextLine[Counter] <> ' ') and
         (TextLine[Counter] <> ',') and
         (TextLine[Counter] <> #09) and
         (TextLine[Counter] <> NewString)) or
         (StringLength=1) then
         TextLength := TextLength+1;


      // Copy first mnemonic and delete it from string
      First := Copy(TextLine, TextStart, TextLength);
      
    end;

    // Delete first word from string
    Delete(TextLine, 1, Counter);

    // Sort out that sodding minus sign
    if MinusCk and (TextLine[1] <> '-') then TextLine := '-' + TextLine

  end
  else
  begin
		// Sort out 0 length string
    First := ' ';
  end;

end;

// *******************************************************************
// Search procedure
//		Starts at the top of file and works down to the bottom of
// 		the file looking for the string
//    {} are comment and default is not search within them
// *******************************************************************
procedure Search;

var	StringDump: String;      // Used as string buffer
	Found,CommentL: Boolean;          // Logic switch to return EOF if search not sucsesful
  i : integer;
  SearchStringU : String;

begin

	// Initialise file pointer and switch

     if (NoReset = False) then Reset(OpenFile);
     SearchStringU := SearchString;
     if (CaseSensitive = False) then SearchStringU := Uppercase(SearchString);
     Found:=false;

 // Loop: Reads in text line, then checks contents.
     repeat
     	Readln(OpenFile, StringDump);
      if (CaseSensitive = False) then StringDump := UpperCase(StringDump);
          if Pos(SearchStringU, StringDump) > 0 then
          	begin
               CommentL := false;
               if NotInComments = true then
                begin;
                 for i := 1 to Pos(SearchStringU, StringDump) do
                    begin;
                     if StringDump[i] = '{' then CommentL := true;
                     if StringDump[i] = '}' then CommentL := false;
                    end;
                end;

               if (CommentL = false) then
                begin;
                Found:=True;
                // Returns string starting at position after the searched string
                Result:= Copy(StringDump, Pos(SearchStringU, StringDump)
					          + Length(SearchStringU), Length(StringDump));
                end;
            end;
     until Found or seekEOF(OpenFile);

     // Return EOF if no string match found
	if Not(Found) then Result:= #26;

end;

{ ******************************************************************* }
{ Procedure SearchSU 										}
{		This is similar to the above search routine except that	}
{		a screen updating function has been incorporated			}
{ ******************************************************************* }
procedure SearchSU;

var	StringDump: String; { Used as string buffer }
	Found: Boolean; { Logic switch to return EOF if search not sucsesful }
     Counter: LongInt; { Screen unpdating Control }

begin

	{ Initialise file pointer and switch }
     Reset(OpenFile);
     Found:=false;
     Counter:= 1;

     { Loop: Reads in text line, then checks contence. }
     repeat
     	Readln(OpenFile, StringDump);
          if ( (Counter Mod 100000) = 0) then Write ('.');
          if Pos(SearchString, StringDump) > 0 then
          	begin
               Found:=True;
               { Returns string starting at position after the searched string }
               Result:= Copy(StringDump, Pos(SearchString, StringDump)
					+ Length(SearchString), Length(StringDump));
               end;
          Counter:= Counter + 1;
     until Found or seekEOF(OpenFile);

     { Return EOF if no string match found }
	if Not(Found) then Result:= #26;

end;


function FindLine(var searchFile: TextFile; searchStr: String;
                  fileReset: Boolean = True; notInComments: Boolean = True;
                  caseSensitive: Boolean = False): String;
{ **************************************************************************** }
// Searches a file for a line containing the requested string and returns the
// found line as a string.
{ **************************************************************************** }

var
  found, commentL: Boolean;
  i: Integer;

begin
  // Read from beginning of file unless optional parameter is false
  if fileReset then Reset(searchFile);
  if caseSensitive = False then searchStr := AnsiUpperCase(searchStr);

  // Search string found flag
  found := False;

  // Search each line of file until line containing search string is found
  while not EoF(searchFile) and not found do
  begin
    ReadLn(searchFile, Result);
    
    if (caseSensitive = False) then Result := AnsiUpperCase(Result);

    if AnsiContainsStr(Result, searchStr) then
    begin
      commentL := False;

      if notInComments then
      begin
        for i := 1 to AnsiPos(searchStr, Result) do
        begin
          if Result[i] = '{' then commentL := True;
          if Result[i] = '}' then commentL := False;
        end;
      end;

      if not commentL then found := True;
    end;
  end;

	// If search string not found
  if EoF(searchFile) and not found then Result := 'Not found';

end;  // FindLine

{ ******************************************************************* }
{ procedure WordDot											}
{	Searchs the first characters in a string looking for '.' This	}
{	is a useful routine when checking file extensions				}
{ ******************************************************************* }
procedure WordDot;

var 	Counter: Integer; { General Loop counting integer }

begin

     { Initialise Varibles etc }
     Counter:= 1;
     IsDot:= 0;

     { Loop until Start of first word is found }
	repeat
  		if (TextLine[Counter] = '.') then IsDot:= 1;
		Counter:= Counter + 1;
     until (Counter = Length(TextLine));

end;

{ ******************************************************************* }
{ SearchBegin precedure (Does NOT RESET FILE)					}
{		Searchs through a file but only checks the first X 		}
{		characters. 										}
{ ******************************************************************* }
procedure SearchBegin;

var	StringDump: String; { Used as string buffer }
	Found: Boolean; { Logic switch to return EOF if search not successful }

begin

	{ Initialise File Switch }
     Found:=false;
     Counter:= 0;

     { Loop: Reads in text line, then checks the begining characters }
     repeat
     	Readln (OpenFile, StringDump);
     	Delete (StringDump, Characters, Length(StringDump));
          if Pos (SearchString, StringDump) > 0 then
          	begin
               Found:=True;
               { Returns string starting at position after the searched string }
               Result:= Copy(StringDump, Pos(SearchString, StringDump)
					+ Length(SearchString), Length(StringDump));
               end;

          Counter:= Counter + 1;
     until Found or seekEOF(OpenFile);

     { Return EOF if no string match found }
	if Not(Found) then Result:= #26;

end;

{ ******************************************************************* }
{ SearchFirstWord                                                   	}
{		Searchs a file but only checks the first numonic or 		}
{		character in each line								}
{ ******************************************************************* }
procedure SearchFirstWord;

var

     Dump,
	StringDump: String; { Used as string buffer }

	Found: Boolean; { Logic switch to return EOF if search not successful }

         isFileName: Boolean;

begin

	{ Initialise file pointer and switch }
     isFileName := false;
     Reset(OpenFile);
     Found:=false;

     { Loop: Reads in text line, then checks contence. }
     repeat
     	Readln(OpenFile, StringDump);
          if (StringDump = '') then
          	begin
			Found:= True;
               Result:= #26;
               end;

          FirstWord (isFileName, StringDump, Dump);
          if (Dump = SearchString) then
          	begin
               Found:=True;

               { Returns string starting at position after the searched string }
               Result:= StringDump;
			end;

     until Found or seekEOF(OpenFile);

     { Return EOF if no string match found }
	if Not(Found) then Result:= #26;

end;

{ ******************************************************************* }
{ Precedure SearchKeyWord                                             }
{		Starts at the top of file and works down to the bottom of   }
{ 		the file. Looks for keywords only, as defined by the		}
{		starting character									}
{ ******************************************************************* }
procedure SearchKeyWord;

var
{     Dump,}
	StringDump: String; { Used as string buffer }

     IsSame: Boolean;			{ Logic switch to tell if the strings are the same }

     Counter: Integer;

begin

	{ Initialise file pointer and switch }
     Reset (OpenFile);
     IsSame:= False;

     { Loop: Reads in text line, then checks contence. }
     repeat
     	begin

          { Read in string and check its not length 0 }
     	Readln(OpenFile, StringDump);
          if (StringDump = '') then
          	begin
			IsSame:= True;
               Result:= #26;
               end

          else if (StringDump[1] = '') or (StringDump = #26) then
          	begin
			IsSame:= True;
               Result:= #26;
               end

          { Check the first characters only }
          else if (StringDump[1] = SearchChar) then
			begin

               { Set the IsSame varible to true, and check to see if its false }
               IsSame:= True;

               { check to see if the string is the same }
               Counter:= 1;
               repeat
                	begin

				{ Check to see if the strings are different }
                    if (SearchString[Counter] <> StringDump[Counter]) then IsSame:= False;

                    { Incriment Counter }
                    Counter:= Counter + 1;
                    end;
               until ((Length (SearchString) +1) = Counter) or (IsSame = False);

			end;
          { Set up loop ending conditions }
          end;
     until IsSame or seekEOF(OpenFile);

     { Return EOF if no string match found }
	if Not(IsSame) then Result:= #26;

end;

{ ******************************************************************* }
{ Precedure SearchNextLine									}
{		From the current possition in the file it returns the 		}
{		next data line, missing comment lines 					}
{ ******************************************************************* }
procedure SearchNextLine;
var

	Dump: String;			// Basic string varible 
     EndLoop: Boolean;		// Loop control
     isFileName: Boolean;

begin

     // Initial varible setup
     isFileName := false;
     EndLoop:= False;

     // Read line and test for a comment line
     repeat
     	begin
		ReadLn (OpenFile, Dump);

          FirstWord (isFileName, Dump, Result);

          if (Result[1] <> 'C') then EndLoop:= True;

          end;
     until (EndLoop = True);

     // Update the result string so it contains the full string
     Result:= Result + ' ' + Dump;

end;

// *******************************************************************
// Precedure SearchToNextNumber
//		Searches file from current position to the the next line
//		starting with a number
// *******************************************************************
procedure SearchToNextNumber;

var

	Dump: String;			// Basic string varible
     Number: Real; 			// number dump
     ErrorCk: Integer;		// Loop control
     isFileName: Boolean;

begin

     // Initial varible setup
     isFileName := false;
     ErrorCk:= 1;

     // Read line and test for a number
     repeat
     	begin
          ReadLn (OpenFile, Dump);
          FirstWord (isFileName, Dump, Result);
          Val (Result, Number, ErrorCk);

          end;

     until (ErrorCk = 0);

     // Update the result string so it contains the full string
     Result:= Result + ' ' + Dump;

end;

// *******************************************************************
// Precedure SearchToNextNonNumber
//		Searches file from current position to the the next line
//		starting with a number
// *******************************************************************
procedure SearchToNextNonNumber;
var

	Dump: String;			// Basic string varible
     Number: Real; 			// number dump
     ErrorCk: Integer;		// Loop control
     isFileName: Boolean;

begin

     // Initial varible setup
     isFileName := False;
     ErrorCk:= 0;

     // Read line and test for a number
     repeat
     	begin
          ReadLn (OpenFile, Dump);
          FirstWord (isFileName, Dump, Result);
          Val (Result, Number, ErrorCk);

          end;

     until (ErrorCk = 1);

     // Update the result string so it contains the full string
     Result:= Result + ' ' + Dump;

end;

// *******************************************************************
// function UpStr
//	Converts a string to Upper Case
// *******************************************************************
function  UpStr;
var
	Counter: Integer;
     // ConStr: String;
begin

	For Counter:= 1 To Length(TextLine) Do
     	begin
		UpStr[Counter]:= UpCase(TextLine[Counter]);
          end;

end;

// *******************************************************************
//  Converts a numonic to a number
// *******************************************************************
function MonthStr2Num (const MonthLine: String): Integer;
var
     LowerString: String;
     Textdump: String;
     isFileName: Boolean;
begin

     isFileName := False;
     TextDump:= LowerCase(MonthLine);
     if length (TextDump) > 0 then
          begin
          if (TextDump[1] = '-') or (TextDump[1] = '/') or (TextDump[1] = '\') or (TextDump[1] = ':') then TextDump[1]:= ' ';
          FirstWord(isFileName, Textdump, LowerString);

          if (LowerString = 'j') or (LowerString = 'jan') or (LowerString = 'january') then MonthStr2Num:= 1
          else if (LowerString = 'f') or (LowerString = 'feb') or (LowerString = 'february') then MonthStr2Num:= 2
          else if (LowerString = 'mch') or (LowerString = 'mar') or (LowerString = 'march') then MonthStr2Num:= 3

          else if (LowerString = 'apr') or (LowerString = 'april') then MonthStr2Num:= 4
          else if (LowerString = 'may') then MonthStr2Num:= 5
          else if (LowerString = 'jne') or (LowerString = 'jun') or (LowerString = 'june') then MonthStr2Num:= 6

          else if (LowerString = 'jul') or (LowerString = 'jly') or (LowerString = 'july') then MonthStr2Num:= 7
          else if (LowerString = 'aug') or (LowerString = 'august') then MonthStr2Num:= 8
          else if (LowerString = 's') or (LowerString = 'sep') or (LowerString = 'sept') or (LowerString = 'september') then MonthStr2Num:= 9

          else if (LowerString = 'o') or (LowerString = 'oct') or (LowerString = 'october') then MonthStr2Num:= 10
          else if (LowerString = 'n') or (LowerString = 'nov') or (LowerString = 'november') then MonthStr2Num:= 11
          else if (LowerString = 'd') or (LowerString = 'dec') or (LowerString = 'december') then MonthStr2Num:= 12
          else MonthStr2Num:= 0;
          end
     else MonthStr2Num:= 0;

end;

end.
