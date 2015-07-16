unit XLSTestRun4;

interface
uses Xc12Utils5, XLSReadWriteII5, XLSNames5, SysUtils, typinfo, Vcl.Dialogs, windows;

procedure XLSCreateTemplateAndLoadData( XLSVersion : TExcelVersion; TemplateFileName, ExportFileName : String ) ;
procedure XLSCreateTemplate( XLSVersion : TExcelVersion; TemplateFileName : String ) ;
procedure XLSPopulateTemplate( XLSVersion : TExcelVersion; TemplateFileName, ExportFileName : String ) ;

implementation

const SHEET1 = 'Summary';
      SHEET2 = 'Data1';
      SHEET3 = 'Data 2';
      SHEET4 = 'Data3';

      SHEET2NAME1 = 'S2NAME1';
      SHEET2NAME2 = 'S2NAME2';

      SHEET3NAME1 = 'S3NAME1';
      SHEET3NAME2 = 'S3NAME2';

      SHEET4NAME1 = 'S4NAME1';
      SHEET4NAME2 = 'S4NAME2';
      SHEET4NAME3 = 'S4NAME3'; //BLANK DATA

procedure XLSPopulateTemplate( XLSVersion : TExcelVersion; TemplateFileName, ExportFileName : String ) ;
var XLSReadWriteII52: TXLSReadWriteII5;
    i : Integer;
    sTemp : String;

  procedure CreateExportFile(iMethodType : Integer);
  begin
    case iMethodType of
      1 : begin
            CopyFile( PChar( TemplateFileName), PChar(ExportFileName), true);
            XLSReadWriteII52.Filename := ExportFileName;
          end
      else begin
        XLSReadWriteII52.Filename := TemplateFileName;
        XLSReadWriteII52.Read;
        XLSReadWriteII52.Filename := ExportFileName;
        XLSReadWriteII52.Write;
      end;
    end;
  end;

begin
  assert( FileExists( TemplateFileName ) , 'File not Created. "'+ XLSReadWriteII52.Filename +'"');



  XLSReadWriteII52 := TXLSReadWriteII5.Create(nil);
  XLSReadWriteII52.Version := XLSVersion;

  CreateExportFile(1);

  try
    assert( FileExists( XLSReadWriteII52.Filename ) , 'File not Created. "'+ XLSReadWriteII52.Filename +'"');
    XLSReadWriteII52.Read;

    for i := 1 to 1 do begin
      if i > 0 then
        XLSReadWriteII52.SheetByName( SHEET2 ).InsertRows(2,1);
      XLSReadWriteII52.SheetByName( SHEET2 ).AsFloat[0, 1+i] := i;
      XLSReadWriteII52.SheetByName( SHEET2 ).AsFloat[1, 1+i] := i mod 2;
      XLSReadWriteII52.SheetByName( SHEET2 ).AsFormula[2, 1+i] := 'SUMIF('+ SHEET2NAME2 + ','+ inttostr( i mod 2 ) +','+ SHEET2NAME1 +')';
      XLSReadWriteII52.Write;  //Try Writing After each insert
    end;

    //Try populating the data in a different cell
    XLSReadWriteII52.SheetByName( SHEET3 ).InsertRows(2,10);
    for i := 10 downto 1 do begin
      XLSReadWriteII52.SheetByName( SHEET3 ).AsFloat[0, 1+i] := i;
      XLSReadWriteII52.SheetByName( SHEET3 ).AsFloat[1, 1+i] := i mod 2;
      XLSReadWriteII52.SheetByName( SHEET3 ).AsFormula[2, 1+i] := 'SUMIF('+  SHEET3NAME2 + ','+ inttostr( i mod 2 ) +','+ SHEET3NAME1 +')';
      if i in [5,6] then
        XLSReadWriteII52.SheetByName( SHEET3 ).AsFormula[3, 1+i] :='SUMIF(' +SHEET3NAME2+ ',INDIRECT( ADDRESS( ROW(),(COLUMN()-2))), '+SHEET3NAME1+')' ;
    end;
    XLSReadWriteII52.Write;

    //Try populating the data in middle of a range
    XLSReadWriteII52.SheetByName( SHEET4 ).InsertRows(2,15);
    for i := 15 downto 1 do begin
      XLSReadWriteII52.SheetByName( SHEET4 ).AsFloat[0, 1+i] := i;
      if (i mod 2) = 0 then
        sTemp := 'A'
      else
        sTemp := 'B';
      XLSReadWriteII52.SheetByName( SHEET4 ).AsString[1, 1+i] := sTemp;
      //XLSReadWriteII52.SheetByName( SHEET4 ).AsFloat[1, 1+i] := i mod 2;

      if i in [5,6] then
        XLSReadWriteII52.SheetByName( SHEET4 ).AsFormula[3, 1+i] :='SUMIF(' +SHEET4NAME2+ ',INDIRECT( ADDRESS( ROW(),(COLUMN()-2))), '+SHEET4NAME1+')';
      XLSReadWriteII52.Write;
    end;

    XLSReadWriteII52.SheetByName( SHEET4 ).InsertRows(6,3);
    for i := 3 downto 1 do begin
      XLSReadWriteII52.SheetByName( SHEET4 ).AsFloat[0, 5+i] := i;
      if (i mod 2) = 0 then
        sTemp := 'A'
      else
        sTemp := 'B';
      XLSReadWriteII52.SheetByName( SHEET4 ).AsString[1, 5+i] := sTemp;
     // XLSReadWriteII52.SheetByName( SHEET4 ).AsFormula[2, 5+i] := 'SUMIF('+ SHEET4NAME2 + ','+ sTemp +','+ SHEET4NAME1 +')';

      if i in [5,6] then
        XLSReadWriteII52.SheetByName( SHEET4 ).AsFormula[3, 5+i] :='SUMIF(' +SHEET4NAME2+ ',INDIRECT( ADDRESS( ROW(),(COLUMN()-2))), '+SHEET4NAME1+')';
      XLSReadWriteII52.Write;
    end;

    XLSReadWriteII52.SheetByName( SHEET4 ).InsertRows(0,5);
    XLSReadWriteII52.SheetByName( SHEET4 ).AsFormula[1, 3] := 'SUM('+#39+SHEET4+#39+'!B1:B3)';

    //XLSReadWriteII52.calculate;
    //XLSReadWriteII52.Write;

    XLSReadWriteII52.SheetByName( SHEET4 ).InsertRows(1,2);
   // XLSReadWriteII52.SheetByName( SHEET4 ).AsFormula[1, 0] := 'SUMIF('+ SHEET4NAME2 + ',A,'+ SHEET4NAME1 +')';
   // XLSReadWriteII52.SheetByName( SHEET4 ).AsFormula[1, 1] := 'SUMIF('+ SHEET4NAME2 + ',B,'+ SHEET4NAME1 +')';
    XLSReadWriteII52.SheetByName( SHEET4 ).AsString[0, 2] := 'A';
    XLSReadWriteII52.SheetByName( SHEET4 ).AsString[0, 3] := 'B';
    XLSReadWriteII52.SheetByName( SHEET4 ).AsFormula[1, 2] :='SUMIF(' +SHEET4NAME2+ ',INDIRECT( ADDRESS( ROW(),(COLUMN()-1))), '+SHEET4NAME1+')';
    XLSReadWriteII52.SheetByName( SHEET4 ).AsFormula[1, 3] :='SUMIF(' +SHEET4NAME2+ ',INDIRECT( ADDRESS( ROW(),(COLUMN()-1))), '+SHEET4NAME1+')';
    XLSReadWriteII52.SheetByName( SHEET4 ).AsFormula[2, 2] :='SUMIF(' +SHEET4NAME2+ ',INDIRECT( ADDRESS( ROW(),(COLUMN()-2))), '+SHEET4NAME3+')';
    XLSReadWriteII52.SheetByName( SHEET4 ).AsFormula[2, 3] :='SUMIF(' +SHEET4NAME2+ ',INDIRECT( ADDRESS( ROW(),(COLUMN()-2))), '+SHEET4NAME3+')';
    XLSReadWriteII52.SheetByName( SHEET4 ).AsFormula[3, 2] :='SUMIF(' +SHEET4NAME3+ ',INDIRECT( ADDRESS( ROW(),(COLUMN()-3))), '+SHEET4NAME1+')';
    XLSReadWriteII52.SheetByName( SHEET4 ).AsFormula[3, 3] :='SUMIF(' +SHEET4NAME3+ ',INDIRECT( ADDRESS( ROW(),(COLUMN()-3))), '+SHEET4NAME1+')';
    XLSReadWriteII52.Write;
    XLSReadWriteII52.calculate;
    XLSReadWriteII52.Write;

  finally
    XLSReadWriteII52.Free;
  end;
end;

procedure XLSCreateTemplate( XLSVersion : TExcelVersion; TemplateFileName : String ) ;
var XLSReadWriteII52: TXLSReadWriteII5;
    i : integer;
begin
  XLSReadWriteII52 := TXLSReadWriteII5.Create(nil);
  XLSReadWriteII52.Version := XLSVersion;
  XLSReadWriteII52.Filename := TemplateFileName;

  try
    XLSReadWriteII52.Write;
    assert( FileExists( XLSReadWriteII52.Filename ) , 'File not Created. "'+ XLSReadWriteII52.Filename +'"');

    XLSReadWriteII52.Sheets[0].Name := SHEET1;
    XLSReadWriteII52.Write;
    XLSReadWriteII52.Add;
    XLSReadWriteII52.Sheets[1].Name := SHEET2;
    XLSReadWriteII52.Write;
    XLSReadWriteII52.Add;
    XLSReadWriteII52.Sheets[2].Name := SHEET3;
    XLSReadWriteII52.Write;
    XLSReadWriteII52.Add;
    XLSReadWriteII52.Sheets[3].Name := SHEET4;
    XLSReadWriteII52.Write;

    XLSReadWriteII52.SheetByName( SHEET1 ).AsString[0,0] := SHEET2 + ' Totals';
    XLSReadWriteII52.SheetByName( SHEET1 ).AsString[3,0] := 'Should Be';
    XLSReadWriteII52.SheetByName( SHEET1 ).AsString[1,1] := 'DATA1';
    XLSReadWriteII52.SheetByName( SHEET1 ).AsFloat[3,1] := 1.3;
    XLSReadWriteII52.SheetByName( SHEET1 ).AsString[1,2] := 'DATA2';
    XLSReadWriteII52.SheetByName( SHEET1 ).AsFloat[3,2] := 1.3;
    XLSReadWriteII52.SheetByName( SHEET1 ).AsString[0,5] := SHEET3 + ' Totals';
    XLSReadWriteII52.SheetByName( SHEET1 ).AsString[1,6] := 'DATA3';
    XLSReadWriteII52.SheetByName( SHEET1 ).AsFloat[3,6] := 55.3;
    XLSReadWriteII52.SheetByName( SHEET1 ).AsString[1,7] := 'DATA4';
    XLSReadWriteII52.SheetByName( SHEET1 ).AsFloat[3,7] := 5.3;
    XLSReadWriteII52.SheetByName( SHEET1 ).AsString[0,10] := SHEET4 + ' Totals';
    XLSReadWriteII52.SheetByName( SHEET1 ).AsFloat[3,11] := 126.6;
    XLSReadWriteII52.SheetByName( SHEET1 ).AsString[1,11] := 'DATA5';
    XLSReadWriteII52.SheetByName( SHEET1 ).AsString[1,12] := 'DATA6';
    XLSReadWriteII52.SheetByName( SHEET1 ).AsFloat[3,12] := 0.6;
    XLSReadWriteII52.Write;

    XLSReadWriteII52.SheetByName( SHEET2 ).AsString[0,0] := 'DATA1';
    XLSReadWriteII52.SheetByName( SHEET2 ).AsString[1,0] := 'DATA2';
    XLSReadWriteII52.names.Add( SHEET2NAME1, ''''+SHEET2+ ''''+'!$A$2:$A$3' );
    XLSReadWriteII52.names.Add( SHEET2NAME2, ''''+SHEET2+ ''''+'!$B$2:$B$3' );
    XLSReadWriteII52.SheetByName( SHEET2 ).AsFloat[0,1] := 0.1;
    XLSReadWriteII52.SheetByName( SHEET2 ).AsFloat[1,1] := 0.1;
    XLSReadWriteII52.SheetByName( SHEET2 ).AsFloat[0,2] := 0.2;
    XLSReadWriteII52.SheetByName( SHEET2 ).AsFloat[1,2] := 0.2;
    XLSReadWriteII52.Write;
    XLSReadWriteII52.SheetByName( SHEET2 ).AsFormula[0,3] := 'sum('+ SHEET2NAME1 +')';
    XLSReadWriteII52.SheetByName( SHEET2 ).AsFormula[1,3] := 'sum('+ SHEET2NAME2 +')';
    XLSReadWriteII52.SheetByName( SHEET2 ).AsFormula[0,4] := 'SUM('+#39+SHEET2+#39+'!A2:A3)';
    XLSReadWriteII52.SheetByName( SHEET2 ).AsFormula[1,4] := 'SUM('+#39+SHEET2+#39+'!B2:B3)';
    XLSReadWriteII52.SheetByName( SHEET2 ).AsString[2,3] := 'Name Sums';
    XLSReadWriteII52.SheetByName( SHEET2 ).AsString[2,4] := 'Sum Sums';

    XLSReadWriteII52.SheetByName( SHEET2 ).Cell[0,3].FillPatternForeColor := xcYellow;
    XLSReadWriteII52.SheetByName( SHEET2 ).Cell[1,3].FillPatternForeColor := xcYellow;
    XLSReadWriteII52.SheetByName( SHEET2 ).Cell[0,4].FillPatternForeColor := xcYellow;
    XLSReadWriteII52.SheetByName( SHEET2 ).Cell[1,4].FillPatternForeColor := xcYellow;
    XLSReadWriteII52.Write;

    XLSReadWriteII52.SheetByName( SHEET3 ).AsString[0,0] := 'DATA3';
    XLSReadWriteII52.SheetByName( SHEET3 ).AsString[1,0] := 'DATA4';
    XLSReadWriteII52.names.Add( SHEET3NAME1, ''''+SHEET3+ ''''+'!$A$2:$A$3' );
    XLSReadWriteII52.names.Add( SHEET3NAME2, ''''+SHEET3+ ''''+'!$B$2:$B$3' );
    XLSReadWriteII52.SheetByName( SHEET3 ).AsFloat[0,1] := 0.1;
    XLSReadWriteII52.SheetByName( SHEET3 ).AsFloat[1,1] := 0.1;
    XLSReadWriteII52.SheetByName( SHEET3 ).AsFloat[0,2] := 0.2;
    XLSReadWriteII52.SheetByName( SHEET3 ).AsFloat[1,2] := 0.2;
    XLSReadWriteII52.Write;
    XLSReadWriteII52.SheetByName( SHEET3 ).AsFormula[0,3] := 'sum('+ SHEET3NAME1 +')';
    XLSReadWriteII52.SheetByName( SHEET3 ).AsFormula[1,3] := 'sum('+ SHEET3NAME2 +')';
    XLSReadWriteII52.SheetByName( SHEET3 ).AsFormula[0,4] := 'SUM('+#39+SHEET3+#39+'!A2:A3)';
    XLSReadWriteII52.SheetByName( SHEET3 ).AsFormula[1,4] := 'SUM('+#39+SHEET3+#39+'!B2:B3)';
    XLSReadWriteII52.SheetByName( SHEET3 ).AsString[2,3] := 'Name Sums';
    XLSReadWriteII52.SheetByName( SHEET3 ).AsString[2,4] := 'Sum Sums';

    XLSReadWriteII52.SheetByName( SHEET3 ).Cell[0,3].FillPatternForeColor := xcYellow;
    XLSReadWriteII52.SheetByName( SHEET3 ).Cell[1,3].FillPatternForeColor := xcYellow;
    XLSReadWriteII52.SheetByName( SHEET3 ).Cell[0,4].FillPatternForeColor := xcYellow;
    XLSReadWriteII52.SheetByName( SHEET3 ).Cell[1,4].FillPatternForeColor := xcYellow;
    XLSReadWriteII52.Write;

    XLSReadWriteII52.SheetByName( SHEET4 ).AsString[0,0] := 'DATA3';
    XLSReadWriteII52.SheetByName( SHEET4 ).AsString[1,0] := 'DATA4';
    XLSReadWriteII52.names.Add( SHEET4NAME1, ''''+SHEET4+ ''''+'!$A$2:$A$4' );
    XLSReadWriteII52.names.Add( SHEET4NAME2, ''''+SHEET4+ ''''+'!$B$2:$B$4' );
    XLSReadWriteII52.names.Add( SHEET4NAME3, ''''+SHEET4+ ''''+'!$F$2:$F$4' );

    XLSReadWriteII52.SheetByName( SHEET4 ).AsFloat[0,1] := 0.1;
    XLSReadWriteII52.SheetByName( SHEET4 ).AsFloat[1,1] := 0.1;
    XLSReadWriteII52.SheetByName( SHEET4 ).AsFloat[0,2] := 0.2;
    XLSReadWriteII52.SheetByName( SHEET4 ).AsFloat[1,2] := 0.2;
    XLSReadWriteII52.SheetByName( SHEET4 ).AsFloat[0,3] := 0.3;
    XLSReadWriteII52.SheetByName( SHEET4 ).AsFloat[1,3] := 0.3;
    XLSReadWriteII52.Write;
    XLSReadWriteII52.SheetByName( SHEET4 ).AsFormula[0,4] := 'sum('+ SHEET4NAME1 +')';
    XLSReadWriteII52.SheetByName( SHEET4 ).AsFormula[1,4] := 'sum('+ SHEET4NAME2 +')';
    XLSReadWriteII52.SheetByName( SHEET4 ).AsFormula[0,5] := 'SUM('+#39+SHEET4+#39+'!A2:A4)';
    XLSReadWriteII52.SheetByName( SHEET4 ).AsFormula[1,5] := 'SUM('+#39+SHEET4+#39+'!B2:B4)';
    XLSReadWriteII52.SheetByName( SHEET4 ).AsString[2,4] := 'Name Sums';
    XLSReadWriteII52.SheetByName( SHEET4 ).AsString[2,5] := 'Sum Sums';
    XLSReadWriteII52.Write;

    XLSReadWriteII52.SheetByName( SHEET4 ).Cell[0,4].FillPatternForeColor := xcYellow;
    XLSReadWriteII52.SheetByName( SHEET4 ).Cell[1,4].FillPatternForeColor := xcYellow;
    XLSReadWriteII52.SheetByName( SHEET4 ).Cell[0,5].FillPatternForeColor := xcYellow;
    XLSReadWriteII52.SheetByName( SHEET4 ).Cell[1,5].FillPatternForeColor := xcYellow;

    XLSReadWriteII52.Write;

    XLSReadWriteII52.SheetByName( SHEET1 ).AsFormula[2,1] := 'sum('+ SHEET2NAME1 +')';
    XLSReadWriteII52.SheetByName( SHEET1 ).AsFormula[2,2] := 'sum('+ SHEET2NAME2 +')';
    XLSReadWriteII52.SheetByName( SHEET1 ).AsFormula[2,6] := 'sum('+ SHEET3NAME1 +')';
    XLSReadWriteII52.SheetByName( SHEET1 ).AsFormula[2,7] := 'sum('+ SHEET3NAME2 +')';
    XLSReadWriteII52.SheetByName( SHEET1 ).AsFormula[2,11] := 'sum('+ SHEET4NAME1 +')';
    XLSReadWriteII52.SheetByName( SHEET1 ).AsFormula[2,12] := 'sum('+ SHEET4NAME2 +')';
    XLSReadWriteII52.Write;
  finally
    XLSReadWriteII52.Free;
  end;
end;

procedure XLSCreateTemplateAndLoadData( XLSVersion : TExcelVersion; TemplateFileName, ExportFileName : String ) ;
begin
  try
    XLSCreateTemplate( XLSVersion, TemplateFileName );
  except
    on E:Exception do begin
      e.Message := 'Error with XLSCreateTemplate fn[XLSTestRun4.XLSCreateTemplateAndLoadData]' + #13#10 +  e.Message;
      raise;
    end;
  end;
  try
    XLSPopulateTemplate( XLSVersion, TemplateFileName, ExportFileName );
  except
    on E:Exception do begin
      e.Message := 'Error with XLSPopulateTemplate fn[XLSTestRun4.XLSCreateTemplateAndLoadData]' + #13#10 +  e.Message;
      raise;
    end;
  end;
end;

end.
