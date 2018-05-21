NET.addAssembly('System.Windows.Forms');
cbdata = System.Windows.Forms.Clipboard.GetDataObject();
memstream = cbdata.GetData('XML Spreadsheet');
bytes = NET.createArray('System.Byte', memstream.Length);
memstream.Read(bytes, int32(0), int32(memstream.Length));
xml = native2unicode(uint8(bytes), 'UTF-8');
xml = strrep(xml, sprintf('\r\n'), sprintf('\n'));
disp(xml);


cbdata = System.Windows.Forms.Clipboard.GetDataObject();
memstream = cbdata.GetData('BIFF12');
bytes = NET.createArray('System.Byte', memstream.Length);
memstream.Read(bytes, int32(0), int32(memstream.Length));
fid = fopen('copied.xlsb', 'wb');
fwrite(fid, uint8(bytes));
fclose(fid);

copyfile copied.xlsb copied.zip
mkdir copied
cd copied
unzip ../copied.zip
cd ..

%{
mac : UTI ÇÕ com.microsoft.excel.xls, 'XLS8'
BIFF8 ÇÃóléqÅH



%}

%{
<?xml version="1.0"?>
<?mso-application progid="Excel.Sheet"?>
<Workbook xmlns="urn:schemas-microsoft-com:office:spreadsheet"
 xmlns:o="urn:schemas-microsoft-com:office:office"
 xmlns:x="urn:schemas-microsoft-com:office:excel"
 xmlns:ss="urn:schemas-microsoft-com:office:spreadsheet"
 xmlns:html="http://www.w3.org/TR/REC-html40">
 <Styles>
  <Style ss:ID="Default" ss:Name="Normal">
   <Alignment ss:Vertical="Bottom"/>
   <Borders/>
   <Font ss:FontName="ÇlÇr ÇoÉSÉVÉbÉN" x:CharSet="128" x:Family="Modern" ss:Size="11"
    ss:Color="#000000"/>
   <Interior/>
   <NumberFormat/>
   <Protection/>
  </Style>
 </Styles>
 <Worksheet ss:Name="Sheet1">
  <Table ss:ExpandedColumnCount="2" ss:ExpandedRowCount="3"
   ss:DefaultColumnWidth="54" ss:DefaultRowHeight="13.5">
   <Row>
    <Cell><Data ss:Type="Number">1</Data></Cell>
   </Row>
   <Row>
    <Cell><Data ss:Type="Boolean">0</Data></Cell>
   </Row>
  </Table>
 </Worksheet>
</Workbook>
%}

formats = cbdata.GetFormats();
for i = 0 : formats.Length - 1
    disp(formats.Get(i));
end
