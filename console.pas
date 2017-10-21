{'*******************************************************************************************
'  Name    : Console.pas
'
'  Author  : Unknown
'  Notice  : Copyright (c) 2009 2H Offshore Engineering Ltd 
'          : All Rights Reserved 
 '
'  Notes   : The main display screen code
'
' 
' CVS VERSION INFO
'
'      Last Checked in by: $Author: harrisor $ 
'      Last Checked in:    $Date: 2009-07-15 10:02:29 $
'              Filename:   $RCSfile: console.pas,v $ 
'              Revision:   $Revision: 1.2 $
'              Tag:        $Name: not supported by cvs2svn $
'******************************************************************************************* }

unit Console;

interface

uses
  Windows;

procedure GetXY(var X, Y: SHORT);
procedure GotoXY(X, Y: SHORT);
function KeyPressed: Boolean;
procedure ClrScr(Attr: Cardinal);
procedure ClrEol;
function ReadKey: Char;
implementation

var
  hIn, hOut: DWORD;

///////////////////////////////////////
// Get the current position of the cursor;
///////////////////////////////////////
procedure GetXY(var X, Y: SHORT);
var
  BuffInfo: TConsoleScreenBufferInfo;
begin
  GetConsoleScreenBufferInfo(hOut, BuffInfo);
  X := BuffInfo.dwCursorPosition.X;
  Y := BuffInfo.dwCursorPosition.Y;
end;

///////////////////////////////////////
// Goto a particular location on screen
///////////////////////////////////////
procedure GotoXY(X, Y: SHORT);
var
  C: TCoord;
begin
  C.X := X;
  C.Y := Y;
  SetConsoleCursorPosition(hOut, C);
end;

////////////////////////////////////////////
// Return true if a key has been pressed
// or a mouse has been moved
////////////////////////////////////////////
function KeyPressed: Boolean;
var
  Buffer: TInputRecord;
  Data: DWORD;
begin
  Result := False;
  PeekConsoleInput(hIn, Buffer, 25, Data);
  if Data <> 0 then
    if Buffer.EventType <> Key_Event then
      FlushConsoleInputBuffer(hIn) // Discard mouse input
    else
      Result := True;
end;

////////////////////////////////////////////
// Read single character without echoing output
////////////////////////////////////////////
function ReadKey: Char;
var
  Ch: Char;
  NumRead: DWORD;
  SaveMode: DWORD;
begin
  GetConsoleMode(hIn, SaveMode);
  SetConsoleMode(hIn, 0);
  NumRead := 0;
  while NumRead < 1 do
    ReadConsole(hIn, @Ch, 1, NumRead, nil);

  SetConsoleMode(hIn, SaveMode);
  Result := Ch;
end;

///////////////////////////////////////
// Clear to the end of the current line
///////////////////////////////////////
procedure ClrEol;
var
  C: TCoord;
  NumWritten: DWord;
begin
  GetXY(C.X, C.Y);
  FillConsoleOutputCharacter(hOut, ' ', 80 - C.X, C, NumWritten);
end;

/////////////////////////////////////////////////////////
// Clear the screen to a particular color
/////////////////////////////////////////////////////////
procedure ClrScr(Attr: Cardinal);
var
  NumWritten: DWord;
  C: TCoord;
  Space: Integer;
begin
  Space := 80 * 25;
  c.X := 0;
  c.Y := 0;

  FillConsoleOutputCharacter(hOut, ' ', Space, c, NumWritten);
  FillConsoleOutputAttribute(hOut, Attr, Space, c, NumWritten);

  GotoXY(0,0);
end;

///////////////////////////////////////
// Write data at a particular location
// in a particular color
///////////////////////////////////////
procedure WriteXY(x, y: Integer; S: string; Attr: Cardinal);
var
  C: TCoord;
  Result: DWORD;
begin
  c.X := x;
  c.Y := y;

  SetConsoleTextAttribute(hOut, Attr);
  SetConsoleCursorPosition(hOut, c);
  WriteConsole(hOut, PChar(S), Length(S), Result, nil);
end;

initialization
  hIn := GetStdHandle(STD_INPUT_HANDLE);
  HOut := GetStdHandle(STD_OUTPUT_HANDLE);
end.
