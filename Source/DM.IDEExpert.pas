unit DM.IDEExpert;

interface

uses
  Winapi.Windows,

  System.SysUtils,
  System.Classes,

  Vcl.Dialogs,
  Vcl.Menus,
  Vcl.Clipbrd,

  ToolsAPI;

type
  TIDEHotKey = class(TNotifierObject, IUnknown, IOTANotifier, IOTAKeyboardBinding)
  strict private
    const
      CR = #13;
      LF = #10;
      CRLF = #13#10;
  private
    procedure InsertClipboardText;
  public
    procedure BindKeyboard(const ABindingServices: IOTAKeyBindingServices);
    function  GetBindingType: TBindingType;
    function  GetDisplayName: string;
    function  GetName: string;
    procedure KeyProc(const AContext: IOTAKeyContext; AKeyCode: TShortCut; var ABindingResult: TKeyBindingResult);
  end;

  procedure Register;

implementation

procedure Register;
begin
  (BorlandIDEServices as IOTAKeyBoardServices).AddKeyboardBinding(TIDEHotKey.Create);
end;

procedure TIDEHotKey.BindKeyboard(const ABindingServices: IOTAKeyBindingServices);
begin
  ABindingServices.AddKeyBinding([ShortCut(Ord('V'), [ssCtrl, ssShift])], KeyProc, nil);
end;

function TIDEHotKey.GetBindingType: TBindingType;
begin
  Result := btPartial;
end;

function TIDEHotKey.GetDisplayName: string;
begin
  Result := '&SQL 붙여넣기';
end;

function TIDEHotKey.GetName: string;
begin
  Result := 'hkPasteFromSQL';
end;

function iif(const ACondition: Boolean; const ATrue, AFalse: string): string;
begin
  Result := AFalse;
  if ACondition then
    Result := ATrue;
end;

procedure TIDEHotKey.InsertClipboardText;
begin
  if not Clipboard.HasFormat(CF_UNICODETEXT) then
  begin
    Exit;
  end;

  var LModule := (BorlandIDEServices as IOTAModuleServices).CurrentModule;
  if not Assigned(LModule) then
  begin
    Exit;
  end;

  var LEditor := LModule.GetCurrentEditor as IOTASourceEditor;
  if not Assigned(LEditor) then
  begin
    Exit;
  end;

  var LEditView := LEditor.GetEditView(0);
  if not Assigned(LEditView) then
  begin
    Exit;
  end;

  var LEditPosition := LEditView.Position;
  var LEditBlock := LEditView.Block;
  var LClipboardText := Clipboard.AsText;

  if LEditBlock.IsValid and (LEditBlock.Size > 0) then
    LEditBlock.Delete;

  // 클립보드에 들어있는 내용을 소스코드로 변경한다
  // 각자의 취향에 따라 소스 포매팅을 해준다
  var LLines := TStringList.Create;
  try
    LLines.Text := LClipboardText;
    var LFormattedText := '';
    for var i := 0 to LLines.Count-1 do
    begin
      LFormattedText :=
        LFormattedText + '''' +
        LLines.Strings[i] + ''' + #13#10' + iif(i = LLines.Count-1, ';', ' +') + CRLF;
    end;

    // 에디터에 값을 넣는다
    LEditPosition.InsertText(LFormattedText);
  finally
    LLines.Free;
  end;
end;

procedure TIDEHotKey.KeyProc(const AContext: IOTAKeyContext; AKeyCode: TShortCut; var ABindingResult: TKeyBindingResult);
var
  LKey: Word;
  LModuleServices: IOTAModuleServices;
  LShift: TShiftState;
  LSourceEditor: IOTASourceEditor;
begin
  ShortCutToKey(AKeyCode, LKey, LShift);

  case LKey of
  Ord('V'):
    begin
      if (LShift = [ssCtrl, ssShift]) and Clipboard.HasFormat(CF_UNICODETEXT) then
      begin
        ABindingResult := krHandled;
        if BorlandIDEServices.QueryInterface(IOTAModuleServices, LModuleServices) = S_OK then
          if LModuleServices.CurrentModule.GetModuleFileEditor(0).QueryInterface(IOTASourceEditor, LSourceEditor) = S_OK then
          begin
            InsertClipboardText;
          end;

        LSourceEditor := nil;
        LModuleServices := nil;
      end;
    end;
  end;
end;

end.
