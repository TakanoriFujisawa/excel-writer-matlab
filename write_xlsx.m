function write_xlsx(filename, data, sheet_name)
% 入力の型，サイズはチェック済みとする
% filename   : "~.xlsx" の形の文字列
% data       : 2次元 or 3次元のセル配列
% sheet_name : data の 3 次元目サイズに対応する文字列セル配列

% 文字列データの切り出し
[celldata, cell_isstr, cell_islogical, sharedStrings] = extract_sharedStrings(data);

%% xlsx ファイルに書き出し
try
    cwd = pwd;
    xlsxdir = tempname;
    % xlsxdir = fillfile(cwd, 'temp_xlsxwrite');
    
    try_rmdir(xlsxdir, 's');
    
    mkdir_p(xlsxdir);
    mkdir_p(xlsxdir, '_rels');
    mkdir_p(xlsxdir, 'xl');
    mkdir_p(xlsxdir, 'xl', '_rels');
    mkdir_p(xlsxdir, 'xl', 'worksheets');

    create_xml_basefiles(xlsxdir, sheet_name);
    create_xml_sharedStrings(xlsxdir, sharedStrings);
    create_xml_sheets(xlsxdir, celldata, cell_isstr, cell_islogical);

    cd(xlsxdir)
    zipfile = [xlsxdir '.zip']; 
    zip(zipfile, '.');
    cd(cwd);
    
    try_movefile(zipfile, filename);
    
    if ~ isempty(regexp(xlsxdir, 'temp_xlsxwrite$', 'once'))
        try_rmdir(xlsxdir);
    end
    
catch ex
    if exist(xlsxdir, 'dir')
        [~, msg] = rmdir(xlsxdir, 's');
        warning(msg); % エラーなしの時は warning は何もしない
    end
    cd(cwd);
    rethrow(ex);
end

end

%%
function [celldata, cell_isstr, cell_islogical, sharedStrings] = extract_sharedStrings(rawcelldata)

cell_isstr = cellfun(@ischar, rawcelldata);
cell_islogical = cellfun(@islogical, rawcelldata);
cell_isnum = cellfun(@isnumeric, rawcelldata);

if ~ all(cell_isstr(:) | cell_isnum(:) | cell_islogical(:))
    error('All cells must be numberic or string');
end

celldata = rawcelldata;
sharedStrings = rawcelldata(cell_isstr);
celldata(cell_isstr) = num2cell(0 : sum(cell_isstr(:)) - 1);

end

%%
function create_xml_sheets(xlsxdir, celldata, cell_isstr)
% 出力形式 ==>
% <!-- number -->
%    <row r="1">
%      <c r="A1"><v>1</v></c>
%    </row>
% <!-- string -->
%    <row r="2">
%      <c r="A2" t="s"><v>0</v></c>
%    </row>
% <!-- logical -->
%    <row r="3">
%      <c r="A3" t="b"><v>0</v></c>
%    </row>
% <!-- formula -->
%    <row r="3">
%      <c r="A3"><f>SUM(C4:E4)</f><v>0</v></c>
%    </row>
%
%
    function refs = create_cellref_table(size_celldata)
        %
        % 参照用に A1 スタイルの以下の表現を作成 (Cell reference)
        % { 'A1', 'B1', ...
        %   'A2', 'B2', ...
        %   'A3', .... }
        
        size_celldata = size_celldata(1 : 2);
        
        colstr = colind_to_name(1 : size_celldata(2));
        colstr = repmat(colstr, size_celldata(1), 1);
        rowstr = cellfun(@num2str, num2cell((1 : size_celldata(1))'), 'UniformOutput', false);
        rowstr = repmat(rowstr, 1, size_celldata(2));
        
        refs = strcat(colstr, rowstr);
        
    end

    function v = format_celldata(v)
        if islogical(v)
            v = double(v);
        end
        
        if isinf(v)
            % xlswrite の仕様に合わせて 65535 (-65535) に変換
            v = num2str(sign(v) * 65535);
        elseif isnan(v)
            v = [];
        else
            v = num2str(v);
        end
    end

cellref_table = create_cellref_table([size(celldata, 1), size(celldata, 2)]);
% 1 ==> '1'
celldata = cellfun(@format_celldata, celldata, 'UniformOutput', false);
celltype = cell(size(celldata));
% 文字列の部分に t="s" を追加するためのセル配列
celltype(cell_isstr) = {' t="s"'};
celltype(cell_islogical) = {' t="b"'};
row_start = cellfun(@num2str, num2cell((1 : size(celldata, 1))'), 'UniformOutput', false);
row_start = strcat('    <row r="', row_start, '">\n');
row_end = repmat({'    </row>\n'}, size(celldata, 1), 1, 1);

for i = 1 : size(celldata, 3)
    sheetfile = fullfile(xlsxdir, 'xl', 'worksheets', sprintf('sheet%d.xml', i));
    fid = fopen(sheetfile, 'w', 'n', 'UTF-8');
    fprintf(fid, '%s\n', ...
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>', ...
        '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">');
    % dimension の計算と出力 (MATLAB のため)
    notemptycell = ~cellfun(@isempty, celldata(:, :, i));
    notemptyrow = any(notemptycell, 2);
    notemptycol = any(notemptycell, 1);
    
    if ~ any(notemptyrow)
        dimstr = 'A1:A1';
    else
        dimstr = [
            cellref_table{find(notemptyrow, 1, 'first'), find(notemptycol, 1, 'first')}, ...
            ':', ...
            cellref_table{find(notemptyrow, 1, 'last'), find(notemptycol, 1, 'last')}];
    end
    fprintf(fid, ['  <dimension ref="', dimstr, '"/>\n']);
    
    % dimension -> sheetFormatPr の順でないと Office がエラーを吐く
    fprintf(fid, '%s\n', '  <sheetFormatPr baseColWidth="12" defaultRowHeight="18"/>');
    
    % sheetData
    fprintf(fid, '  <sheetData>\n');
    
    % 文字列                   <c r="         A1        "         t="s"         ><v>          1            </v></c>
    % 数字                     <c r="         A1        "                       ><v>          1            </v></c>
    sheet_data = strcat('      <c r="', cellref_table, '"', celltype(:, :, i), '><v>', celldata(:, :, i), '</v></c>\n');
    sheet_data(~ notemptycell) = {''};
    sheet_data = [row_start, sheet_data, row_end];
    sheet_data = sheet_data(notemptyrow, :);
    sheet_data = sheet_data';
    
    if ~ isempty(sheet_data)
        fprintf(fid, strcat(sheet_data{:}));
    end
    
    fprintf(fid, '  </sheetData>\n</worksheet>\n');
    fclose(fid);
end

end


%%
function str = xml_escape_string(str)
% xml 向けにエスケープ
% str : 文字列 / 文字列セル

if ischar(str)
    str = char(org.apache.commons.lang.StringEscapeUtils.escapeXml(str));
else
    for i = 1 : length(str)
        str{i} = char(org.apache.commons.lang.StringEscapeUtils.escapeXml(str{i}));
    end
end
% ↑ は日本語をUnicodeにエスケープするので不都合なら↓を

%xml_reserved = { '&', '"', '''', '<', '>'};
%xml_escaped  = { '&amp;', '&quot;', '&apos;', '&lt;', '&gt;'};

%for i = 1 : length(xml_reserved)
%    sharedStrings = strrep(sharedStrings, xml_reserved{i}, xml_escaped{i});
%end

end

%%
function create_xml_sharedStrings(xlsxdir, sharedStrings)

fid = fopen(fullfile(xlsxdir, 'xl', 'sharedStrings.xml'), 'w', 'n', 'UTF-8');

fprintf(fid, '%s\n',  ...
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>', ...
    '<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">');

if ~ isempty(sharedStrings)
    % xml 向けにエスケープ
    sharedStrings = xml_escape_string(sharedStrings);
    fprintf(fid, '  <si><t>%s</t></si>\n', sharedStrings{:});
end

fprintf(fid, '</sst>');

fclose(fid);

end

%%
function create_xml_basefiles(xlsxdir, sheet_name)

n_sheets = length(sheet_name);

% 'ＭＳ Ｐゴシック'
font_ms_pgothic = char([65325 65331 32 65328 12468 12471 12483 12463]);
% '標準'
style_normal = char([27161 28310]);

sheet_name = xml_escape_string(sheet_name);
% シート名, シートの番号 2 回
sheet_name_id_rid = cell(3, n_sheets);
for n = 1 : n_sheets
    sheet_name_id_rid{:, n} = { sheet_name{n}; n; n };
end

xml_contents = {
    '[Content_Types].xml', {
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
        '  <Default Extension="xml" ContentType="application/xml"/>'
        '  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
        '  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>'
        sprintf('  <Override PartName="/xl/worksheets/sheet%d.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>\n', 1 : n_sheets);
        '  <Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>'
        '  <Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/>'
        '</Types>'
    }
    
    '_rels/.rels', {
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        '  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>'
        '</Relationships>'
    }
    
    'xl/workbook.xml', {
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"'
        '	  xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
        '  <sheets>'
        sprintf('    <sheet name="%s" sheetId="%d" r:id="rId%d"/>\n', sheet_name_id_rid{:});
        '  </sheets>'
        '</workbook>'
    }
    
    'xl/_rels/workbook.xml.rels', {
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        sprintf('  <Relationship Id="rId%d" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet%d.xml"/>\n', [1; 1] * (1 : n_sheets));
        sprintf('  <Relationship Id="rId%d" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>\n', n_sheets + 1);
        sprintf('  <Relationship Id="rId%d" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" Target="sharedStrings.xml"/>\n', n_sheets + 2);
        '</Relationships>'
    }

    'xl/styles.xml', {
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
        '  <fonts count="2">'
        ['    <font><sz val="12"/><name val="' font_ms_pgothic '"/><family val="2"/><charset val="128"/><scheme val="minor"/></font>']
        ['    <font><sz val="6"/><name val="' font_ms_pgothic '"/><family val="2"/><charset val="128"/><scheme val="minor"/></font>']
        '  </fonts>'
        '  <fills count="2">'
        '    <fill><patternFill patternType="none"/></fill>'
        '    <fill><patternFill patternType="gray125"/></fill>'
        '  </fills>'
        '  <borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>'
        '  <cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>'
        '  <cellXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/></cellXfs>'
        ['  <cellStyles count="1"><cellStyle name="' style_normal '" xfId="0" builtinId="0"/></cellStyles>']
        '  <dxfs count="0"/>'
        '  <tableStyles count="0" defaultTableStyle="TableStyleMedium9" defaultPivotStyle="PivotStyleMedium4"/>'
        '</styleSheet>'
    }
};

for x = 1 : size(xml_contents, 1)
    fid = fopen(fullfile(xlsxdir, xml_contents{x, 1}), 'w', 'n', 'UTF-8');
    contents = regexprep(xml_contents{x, 2}, '\n$', '');
    fprintf(fid, '%s\n', contents{:});
    fclose(fid);
end

end

%%
