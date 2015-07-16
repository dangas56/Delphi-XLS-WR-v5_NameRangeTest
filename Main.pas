unit Main;

interface

uses
  Winapi.Windows, Winapi.Messages, System.SysUtils, System.Variants, System.Classes, Vcl.Graphics,
  Vcl.Controls, Vcl.Forms, Vcl.Dialogs, Vcl.StdCtrls;

type
  TForm10 = class(TForm)
    btnNameRange: TButton;
    procedure btnNameRangeClick(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  Form10: TForm10;

implementation
uses XLSTestRun4, ShellAPI, XLSNames5, Xc12Utils5, XLSUtils5 ;
{$R *.dfm}

procedure TForm10.btnNameRangeClick(Sender: TObject);
var sTemplateLocation,
    sExportlocation,
    sDateTimeStamp : String;
begin
  sDateTimeStamp := FormatDateTime('yymmdd_hhnnss_zzz', now);
  try
    sTemplateLocation := IncludeTrailingPathDelimiter( ExtractFilePath(Application.exename) ) +
                                            'XLSTestV97TEMP' + sDateTimeStamp + '.XLS';
    sExportLocation   := IncludeTrailingPathDelimiter( ExtractFilePath(Application.exename) ) +
                                            'XLSTestV97EXP' + sDateTimeStamp + '.XLS';
    XLSCreateTemplateAndLoadData(xvExcel97,sTemplateLocation, sExportLocation );

    ShellExecute(self.Handle, 'Open', PChar(sExportLocation), PChar(''), nil, 1);
  except
    on E:Exception Do begin
      e.Message := 'Error when exporting XLS Verion xvExcel97' + #13#10 + e.Message;
      //raise;
      messagedlg(e.Message, mtError, [mbOk], 0);
    end;
  end;

  try
    sTemplateLocation := IncludeTrailingPathDelimiter( ExtractFilePath(Application.exename) ) +
                                            'XLSTestV2007TEMP' + sDateTimeStamp + '.XLSX' ;
    sExportLocation := IncludeTrailingPathDelimiter( ExtractFilePath(Application.exename) ) +
                                            'XLSTestV2007EXP' + sDateTimeStamp + '.XLSX' ;
    XLSCreateTemplateAndLoadData(xvExcel2007, sTemplateLocation, sExportLocation );

    ShellExecute(self.Handle, 'Open', PChar(sExportLocation), PChar(''), nil, 1);
  except
    on E:Exception Do begin
      e.Message := 'Error when exporting XLS Verion xvExcel2007' + #13#10 + e.Message;
      //raise;
      messagedlg(e.Message, mtError, [mbOk], 0);
    end;
  end;
end;

end.
