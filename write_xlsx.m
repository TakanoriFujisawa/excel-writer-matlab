function write_xlsx(filename, data, sheet_name)
% ���͂̌^�C�T�C�Y�̓`�F�b�N�ς݂Ƃ���
% filename   : "~.xlsx" �̌`�̕�����
% data       : 2���� or 3�����̃Z���z��
% sheet_name : data �� 3 �����ڃT�C�Y�ɑΉ����镶����Z���z��

% ������f�[�^�̐؂�o��
[celldata, cell_isstr, cell_islogical, sharedStrings] = extract_sharedStrings(data);

%% xlsx �t�@�C���ɏ����o��
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
        warning(msg); % �G���[�Ȃ��̎��� warning �͉������Ȃ�
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
% �o�͌`�� ==>
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
        % �Q�Ɨp�� A1 �X�^�C���̈ȉ��̕\�����쐬 (Cell reference)
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
            % xlswrite �̎d�l�ɍ��킹�� 65535 (-65535) �ɕϊ�
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
% ������̕����� t="s" ��ǉ����邽�߂̃Z���z��
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
    % dimension �̌v�Z�Əo�� (MATLAB �̂���)
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
    
    % dimension -> sheetFormatPr �̏��łȂ��� Office ���G���[��f��
    fprintf(fid, '%s\n', '  <sheetFormatPr baseColWidth="12" defaultRowHeight="18"/>');
    
    % sheetData
    fprintf(fid, '  <sheetData>\n');
    
    % ������                   <c r="         A1        "         t="s"         ><v>          1            </v></c>
    % ����                     <c r="         A1        "                       ><v>          1            </v></c>
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
% xml �����ɃG�X�P�[�v
% str : ������ / ������Z��

if ischar(str)
    str = char(org.apache.commons.lang.StringEscapeUtils.escapeXml(str));
else
    for i = 1 : length(str)
        str{i} = char(org.apache.commons.lang.StringEscapeUtils.escapeXml(str{i}));
    end
end
% �� �͓��{���Unicode�ɃG�X�P�[�v����̂ŕs�s���Ȃ火��

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
    % xml �����ɃG�X�P�[�v
    sharedStrings = xml_escape_string(sharedStrings);
    fprintf(fid, '  <si><t>%s</t></si>\n', sharedStrings{:});
end

fprintf(fid, '</sst>');

fclose(fid);

end

%%
function create_xml_basefiles(xlsxdir, sheet_name)

n_sheets = length(sheet_name);

% '�l�r �o�S�V�b�N'
font_ms_pgothic = char([65325 65331 32 65328 12468 12471 12483 12463]);
% '�W��'
style_normal = char([27161 28310]);

sheet_name = xml_escape_string(sheet_name);
% �V�[�g��, �V�[�g�̔ԍ� 2 ��
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
