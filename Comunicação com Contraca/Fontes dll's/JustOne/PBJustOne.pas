{ -------------------------------------------------------------------------------------}
{ A "JustOne" component for Delphi32.                                                  }
{ Copyright 1997, Patrick Brisacier.  All Rights Reserved.                             }
{ This component can be freely used and distributed in commercial and private          }
{ environments, provided this notice is not modified in any way.                       }
{ -------------------------------------------------------------------------------------}
{ Feel free to contact me if you have any questions, comments or suggestions at        }
{   PBrisacier@mail.dotcom.fr (Patrick Brisacier)                                      }
{ You can always find the latest version of this component at:                         }
{   http://www.worldnet.net/~cycocrew/delphi/                                          }
{ -------------------------------------------------------------------------------------}
{ Date last modified:  04/06/97                                                        }
{ -------------------------------------------------------------------------------------}

{ -------------------------------------------------------------------------------------}
{ TPBJustOne v1.01                                                                     }
{ -------------------------------------------------------------------------------------}
{ Description:                                                                         }
{   A component that enables only one unique instance of an application at each time.  }
{ -------------------------------------------------------------------------------------}
{ Revision History:                                                                    }
{ 1.00:  + Initial release                                                             }
{ 1.01:  + Cleaned source code                                                         }
{ -------------------------------------------------------------------------------------}
{ Note:                                                                                }
{ There is a limitation under Windows NT 4.0 : when the application is called again,   }
{ it doesn't come to the front. It's not the case under Windows 95.                    }
{ -------------------------------------------------------------------------------------}

unit PBJustOne;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs;

type
  TPBJustOne = class(TComponent)
  private
    { Déclarations privées }
  protected
    { Déclarations protégées }
  public
    constructor Create(AOwner: TComponent); override;
  published
    { Déclarations publiées }
  end;

procedure Register;

implementation

const
  AllowedInstances = 1;

var
  MyAppName, MyClassName: array[0..255] of Char;
  NumFound: Integer;
  LastFound, MyPopup: HWND;

function LookAtAllWindows(Handle: HWND; Temp: LongInt): BOOL; stdcall;
var
  WindowName, ClassName: Array[0..255] of Char;
begin
  // Go get the windows class name
  if GetClassName(Handle, ClassName, SizeOf(ClassName)) > 0 then
    // Is the window class the same ?
    if StrComp(ClassName, MyClassName) = 0 then
      // Get its window caption
      if GetWindowText(Handle, WindowName, SizeOf(WindowName)) > 0 then
        // Does this have the same window title ?
        if StrComp(WindowName, MyAppName) = 0 then
          begin
            inc(NumFound);
            // Are the handles different ?
            if Handle <> Application.Handle then
              // Save it so we can bring it to the top later.
              LastFound := Handle;
          end;

  result := true;
end;

procedure Register;
begin
  RegisterComponents('Système', [TPBJustOne]);
end;

constructor TPBJustOne.Create(AOwner: TComponent);
begin
  inherited;
  NumFound := 0; LastFound := 0;
  // First, determine what this application'name is.
  GetWindowText(Application.Handle, MyAppName, SizeOf(MyAppName));
  // Now determine the class name for this application
  GetClassName(Application.Handle, MyClassName, SizeOf(MyClassName));
  // Now count how many others out there are Delphi apps with this title
  EnumWindows(@LookAtAllWindows, 0);
  if NumFound > AllowedInstances then
    // there is another instance running, bring it to the front !
    begin
      MyPopup := GetLastActivePopup(LastFound);
      // Bring it to the top in the Z-Order
      BringWindowToTop(LastFound);
      // Is the window iconized ?
      if IsIconic(MyPopup) then
        // Restore it to its original position
        ShowWindow(MyPopup, SW_RESTORE)
      else
        // Bring it to the front
        SetForegroundWindow(MyPopup);
      // Halt this instance
      Halt;
    end
end;

end.
