classdef XLSX < handle
    properties
        data = {}
        sheet_name = {}
        show_table = false
        hdl_table = []
    end
    
    methods
        % オーバーロードすべきメソッド
        %
        % subsref, subasgn
        % end
        % numel, size
        % double, char (型変換)
        %
        
        %%
        % Property : sheet_name
        %
        function ret = get_sheet_name(this, subs)
            for i = length(this.sheet_name) + 1 : size(this.data, 3)
                this.sheet_name{i} = [ 'Sheet' num2str(i) ]; 
            end
            
            if nargin == 1 || isempty(subs)
                ret = this.sheet_name;
            else
                ret = subsref(this.sheet_name, subs(1));
            end
        end
        
        function name = get_sheet_name_from_index(this, ind)
            % ind 番目のシートの名前を返す，設定されていなければ
            % デフォルトの名前を返す
            %
            %  % スカラー入力時
            %  xl.get_sheet_name_from_index(1)     => 'Sheet1'
            %
            %  % ベクトル入力時
            %  xl.get_sheet_name_from_index(2 : 3) => {'Sheet2', 'Sheet3'}
            validateattributes(ind, {'numeric'}, {'positive', 'integer'});
            
            if any(ind > size(this.data, 3))
                error('インデックスがシートの範囲外です');
            end
            
            if isscalar(ind)
                name = cell(1, 1);
            else
                name = cell(1, numel(ind));
            end
            
            for i = 1 : length(ind)
                if ind(i) > length(this.sheet_name) || isempty(this.sheet_name{ind(i)})
                    name{i} = ['Sheet' num2str(ind(i))];
                else
                    name{i} = this.sheet_name{ind(i)};
                end
            end
            
            if isscalar(ind)
                name = name{1};
            end
        end
        
        function this = set_sheet_name(this, val, subs)
            if ischar(val)
                validateattributes(val, {'char'}, {'row'});
                val = { val };
                
            elseif iscell(val)
                val = val(:)';
                for i = 1 : length(val)
                    validateattributes(val{i}, {'char'}, {'row'});
                end
            else
                error('シート名は文字列または文字列のセル配列が必要です');
            end
            
            if nargin == 2 || isempty(subs)
                this.sheet_name = val;
            else
                subs(1).type = '()';
                this.sheet_name = subsasgn(this.sheet_name, subs(1), val);
            end
            
            if this.get_show_table()
                this.hdl_table = update_xlsx_table(this.hdl_table, this);
            end
        end
        
        %%
        % Property : show_table
        %
        function ret = get_show_table(this, ~)
            if isempty(this.hdl_table) || ~ ishandle(this.hdl_table)
                this.show_table = false;
            end
            ret = this.show_table;
        end
        
        function this = set_show_table(this, val, ~)
            if val
                this.show_table = true;
                if isempty(this.hdl_table) || ~ ishandle(this.hdl_table)
                    this.hdl_table = setup_xlsx_table();
                else
                    this.hdl_table.Visible = 'on';
                end
                this.hdl_table = update_xlsx_table(this.hdl_table, this);
            else
                this.show_table = false;
                if ishandle(this.hdl_table)
                    this.hdl_table.Visible = 'off';
                end
            end
        end
        
        %%
        % Property : data
        %
        function ret = get_data(this, subs)
            if nargin == 1 || isempty(subs)
                ret = this.data;
            else
                ret = subsref(this.data, subs);
            end
        end
        
        function this = set_data(this, val, subs)
            % subs : subsasgn の S (struct)
            % { [1 : 2], ':', 2 }
            % { [3x3 logical] }
            %
            % 動作例 :
            % x(1, 2 : 3) = 'str'
            % x(2, 4) = 1;
            % x(2, 4 : 6) = 1 : 3;
            %
            %
            % x(3, 2) = rand(5);
            % x(4, 2) = {'str', 2, true}
            %  [この書式の場合インデックスはスカラーでなければいけない]
            %
            
            if isempty(val)
                val = { val };
                
            elseif ischar(val)
                val = { val(:)' };
                
            elseif isnumeric(val) || islogical(val)
                if ndims(val) > 3
                    error('入力は最大３次元の配列である必要があります');
                end
                val = num2cell(val);
                
            elseif iscell(val)
                if ndims(val) > 3
                    error('入力は最大３次元の配列である必要があります');
                end
                
                if ~ all(cellfun(@isxlsxdata, val))
                    error('セルの要素は数値，論理値，または文字列である必要があります');
                end
                
            else
                error('入力は文字列・数値配列またはセル配列が必要です');
            end
            
            if nargin == 3 && ~ isempty(subs)
                % cell に変換して subsasgn に渡す
                %this.data = subsasgn(this.data, subs, val);
                switch subs.type
                    case {'()', '{}'}
                        switch length(subs.subs)
                            case {1, 2}
                                % 添字代入のルールが複雑すぎてよく分からなかったので
                                % 一旦 1 次元スライスに対して subsasgn をした後合成
                                sz = [size(this.data, 1), size(this.data, 2), size(this.data, 3)];
                                tmp_data = subsasgn(this.data(:, :, 1), subs, val);
                                tmp_data(1 : sz(1), 1 : sz(2), 2 : sz(3)) = this.data(:, :, 2 : end);
                                
                            case 3
                                tmp_data = subsasgn(this.data, subs, val);
                                
                            otherwise
                                error(message('MATLAB:badsubscriptTextDimension'));
                        end
                        
                        % 4 次元配列以上にならないよう一旦 tmp_data に代入して
                        % エラーチェック
                        if ndims(tmp_data) > 3
                            error(message('MATLAB:badsubscriptTextDimension'));
                        end
                        
                        this.data = tmp_data;
                        
                    otherwise
                        error(message('MATLAB:nonStrucReference'));
                end
            else
                this.data = val;
            end
            
            if this.get_show_table()
                this.hdl_table = update_xlsx_table(this.hdl_table, this);
            end
        end

        
        %%
        % Constructor
        %
        function this = XLSX(varargin)
            % x = XLSX();
            % x = XLSX(numeric_data);
            % x = XLSX(cell_data);
            % x = XLSX(data, {'Sheet1', 'Sheet2', 'Sheet3'})
            
            % ファイルから (データの損失に注意！)
            % x = XLSX(xls_filename);
            
            % "for internal use only"
            % x = XLSX(data, sheet_name, 'valid')
            
            if nargin == 0
                return
            end
            
            if nargin == 3 && strcmp(varargin{3}, 'valid')
                this.data = varargin{1};
                this.sheet_name = varargin{2};
                return
            end
            
            if nargin == 1 && ischar(varargin{1})
                % Read from xlsx file
                [~, ~, ext] = fileparts(varargin{1});
                if ~ strcmp(ext, '.xlsx')
                    error('XLSX 形式以外の読み取りはサポートしていません');
                end
                
                [this.data, this.sheet_name] = read_xlsx(varargin{1});
                return
            end
            
            if isempty(varargin{1})
                % do nothing
                
            elseif isnumeric(varargin{1}) || islogical(varargin{1}) || iscell(varargin{1})
                this.set_data(varargin{1});
                
            else
                error('第１引数は数値配列・論理値配列，またはセル配列を指定します');
                
            end
            
            if nargin == 2
                this.set_sheet_name(varargin{2});
            end
        end
        
        %%
        % Overloaded Methods
        %
        function disp(this)
            if isempty(this.data)
                fprintf('Empty Sheet\n');
            else
                for ind = 1 : size(this.data, 3)
                    fprintf('<<%d: %s>>\n', ind, this.get_sheet_name_from_index(ind));
                    disp(this.data(:, :, ind));
                end
            end
        end
        
        function ret = subsref(this, subs)
            function check_function_call(s)
                switch s.type
                    case '{}'
                        error(message('MATLAB:cellRefFromNonCell'));
                    case '.'
                        error(message('MATLAB:nonStrucReference'));
                end
            end
            
            switch subs(1).type
                case '()'
                    ret = this.get_data(subs);
                    
                case '{}'
                    error(message('MATLAB:cellRefFromNonCell'));
                    
                case '.'
                    switch subs(1).subs
                        case 'sheet_name'
                            ret = this.get_sheet_name(subs(2 : end));
                            
                        case 'show_table'
                            ret = this.get_show_table(subs(2 : end));
                            
                        case 'data'
                            ret = this.get_data(subs(2 : end));
                            
                        case 'hdl_table'
                            ret = this.hdl_table;

                        case 'set_sheet_name'
                            check_function_call(subs(2));
                            ret = this.set_sheet_name(subs(2).subs{:});
                            
                        case 'get_sheet_name_from_index'
                            check_function_call(subs(2));
                            ret = this.get_sheet_name_from_index(subs(2).subs{:});
                            
                        otherwise
                            error('不明なプロパティ %s への参照です', subs(1).subs);
                    end
            end
        end
        
        function this = subsasgn(this, subs, val)
            switch subs(1).type
                case '()'
                    this.set_data(val, subs(1));
                    
                case '{}'
                    error(message('MATLAB:cellRefFromNonCell'));
                    
                case '.'
                    switch subs(1).subs
                        case 'sheet_name'
                            this.set_sheet_name(val, subs(2 : end));
                            
                        case 'show_table'
                            this.set_show_table(val, subs(2 : end));
                    end
            end
        end
        
        function ind = end(this, the_indexed, total_indexed)
            switch total_indexed
                case 1
                    % x(end) = 2
                    % データが他のシートに入っちゃうことがあるのであまり推奨しない
                    ind = numel(this.data);
                    
                case 2
                    % x(1, end) = 3
                    % 3次元配列の場合と違って第1セルのインデックスを返す
                    ind = size(this.data, the_indexed);
                case 3
                    % x(1:end, 1:end, 1) = 1
                    ind = size(this.data, the_indexed);
                    
                otherwise
                    error('添字の数が多すぎます')
            end
        end
        
        
        
        
        %%
        function write(this, outfile)
            if isempty(outfile)
                error(message('MATLAB:xlswrite:EmptyFileName'));
            elseif ~ ischar(outfile) || ~ isrow(outfile)
                error(message('MATLAB:xlswrite:InputClassFilename'));
            elseif ~ isempty(regexp(outfile, '[*?|^><]', 'once'))
                error(message('MATLAB:xlswrite:FileName'));
            end
            
            % 拡張子を追加
            [directory, outfile, ext] = fileparts(outfile);
            if isempty(ext) || ~ strcmpi(ext, '.xlsx')
                outfile = fullfile(directory, [outfile, ext, '.xlsx']);
            else
                outfile = fullfile(directory, [outfile, ext]);
            end
            
            %% 出力セルのチェック
            if isempty(this.data)
                error(message('MATLAB:xlswrite:EmptyInput'));
            end

            write_xlsx(outfile, this.data, this.get_sheet_name());
        end
        
    end
    
end

%%


%%
function name = colind_to_name(ind)
% 列番号をアルファベット表記に変換 1 -> A
% 入力はスカラー自然数のみ

if nargin == 0
    do_check = @(i)fprintf('%5d : %s\n', i, colind_to_name(i));
    do_check(1);
    do_check(26);    % Z
    do_check(27);    % AA
    do_check(52);    % AZ
    do_check(53);    % BA
    do_check(78);    % BZ
    do_check(79);    % CA
    do_check(702);   % ZZ
    do_check(703);   % AAA
    do_check(16384); % XFD    
    return
    
elseif ~ isscalar(ind)
    name = cellfun(@colind_to_name, num2cell(ind), 'UniformOutput', false);
    return
    
end

if ind < 27
    name = char('A' + ind - 1);
    
elseif ind < 703
    % [A=0 ~ Z=25 までの 26 進数 2 桁] + 27
    ind = ind - 27;
    name = char('A' + [ floor(ind / 26), mod(ind, 26) ]);
    
elseif ind < 17575  % 一応 16384(XFD) までらしい
    % [A=0 ~ Z=25 までの 26 進数 3 桁] + 703
    ind = ind - 703;
    name = char('A' + [ floor(ind / 676), mod(floor(ind / 26), 26), mod(ind, 26) ]);
    
else
    error('Row index exceeded the capacity.');
end

end



%%
function tf = isxlsxdata(v)

if isempty(v)
    tf = true;
    
elseif isnumeric(v) || islogical(v)
    tf = numel(v) == 1;
    
elseif ischar(v)
    tf = isrow(v);
    
else
    tf = false;
    
end

end



%%
function hdl = setup_xlsx_table(hdl, hdl_xlsx)
%
% Figure
%   +--Children=TabGroup
%   |             +--Children(1)=Tab
%   |             |     +--Children=Table
%   |             |     |     +--Data={...}
%   |             |     |     +--ColumnName={...}
%   |             |     +--Title='Sheet1'
%   |             +--Children(2)=Tab
%   |             |     +--Children=Table
%   |             |     +--Title='Sheet2'
%   |             +--Children(3)=Tab
%   |                   +--Children=Table
%   |                   +--Title='Sheet3'
%   +--Visible='on'
%   +--UserData=XLSX
%
% データ追加
%  1. シートの数を調べる
%  2. タブが足りない場合はタブ・テーブルを追加
%  3. タブのTitleにシート名をセットする
%  4. テーブルに CellEditCallback をセット
%  5. タブにデータを追加 (データを分配)
%
% タブにデータを追加
%  1. 必要数のColumnNameを生成して登録
%  2. 行数だけの { 'char' } セル配列を生成して ColumnFormat にセット
% [3. データを文字列に変換 (編集を可能にする場合)]
%  4. データを最低 20 行 6 列になるよう拡張 (見栄えのため)
%  5. データをセット
%

if length(dbstack) == 1
    hdl = gcf;
end

if ~ exist('hdl', 'var') || isempty(hdl)
    hdl = figure('IntegerHandle', 'off', 'ToolBar', 'none');
else
    clf(hdl, 'reset');
end

hgrp = uitabgroup('Parent', hdl, 'TabLocation', 'bottom');

for t = 1 : 3
    htab = uitab('Parent', hgrp, 'Title', ['Sheet' num2str(t)]);
    uitable('Parent', htab, 'Units', 'normalized', 'Position', [0 0 1 1], ...
        'RowName', 'numbered', 'ColumnName', colind_to_name(1 : 6), ...
        'Data', cell(20, 6));
end

if exist('hdl_xlsx', 'var')
    hdl.UserData = hdl_xlsx;
    hdl.DeleteFcn = @(src, ev) hfig_DeleteFcn(src, ev);
end

end

function hfig_DeleteFcn(src, ~)

if ~ isempty(src.UserData)
    src.UserData.show_table = false;
    src.UserData = [];
end

end




%%
function hdl = update_xlsx_table(hdl, xlsx)
% 
% hdl = update_xlsx_table(hdl, xlsx)
%
% Figure
%   +--Children=TabGroup
%   |             +--Children(1)=Tab
%   |             |     +--Children=Table
%   |             |     |     +--Data={...}
%   |             |     |     +--ColumnName={...}
%   |             |     +--Title='Sheet1'
%
% データ追加
%  1. シートの数を調べる
%  2. タブが足りない場合はタブ・テーブルを追加
%  3. タブのTitleにシート名をセットする
%  4. テーブルに CellEditCallback をセット
%  5. タブにデータを追加 (データを分配)
%
% タブにデータを追加
%  1. 必要数のColumnNameを生成して登録
%  2. 行数だけの { 'char' } セル配列を生成して ColumnFormat にセット
% [3. データを文字列に変換 (編集を可能にする場合)]
%  4. データを最低 20 行 6 列になるよう拡張 (見栄えのため)
%  5. データをセット
%

if nargout == 0
    warning('[Internal] 更新後のハンドルを返り値で受け取る必要があります');
end

if ishandle(hdl) && strcmp(hdl.Visible, 'off')
    return
elseif ~ ishandle(hdl)
    % Figure が閉じられている時
    hdl = setup_xlsx_table([], xlsx);
end


for t = 1 : size(xlsx.data, 3)
    if t > length(hdl.Children.Children)
        % 足りないタブを追加
        htab = uitab('Parent', hdl.Children, 'Title', ['Sheet', num2str(t)]);
        htbl = uitable('Parent', htab, 'Units', 'normalized', 'Position', [0 0 1 1], 'RowName', 'numbered');
    else
        htab = hdl.Children.Children(t);
        htbl = htab.Children(1);
    end
    
    htab.Title = xlsx.get_sheet_name_from_index(t);
    
    % 表示用のデータの取得
    data = xlsx.data(:, :, t);
    
    % 見栄えのためサイズを 20x6 まで拡張
    if size(data, 1) < 20
        data(end + 1 : 20, :) = {''};
    end
    
    if size(data, 2) < 6
        data(:, end + 1 : 6) = {''};
    end
    
    % ColumnFormat と ColumnName を必要数用意
    if ~ iscell(htbl.ColumnFormat) || length(htbl.ColumnFormat) < size(data, 2)
        htbl.ColumnFormat = repmat({'char'}, 1, size(data, 2) * 2);
    end
    
    if ~ iscell(htbl.ColumnName) || length(htbl.ColumnName) < size(data, 2)
        htbl.ColumnName = colind_to_name(1 : size(data, 2) + 3);
    end
    
    % データをセット
    htbl.Data = data;
end

end


%%
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
%      <c r="A3"><f></f><v>0</v></c>
%    </row>
%
%

cellref_table = create_cellref_table([size(celldata, 1), size(celldata, 2)]);


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
% 参照用に A1 スタイルの以下の表現を作成 (Cell reference)
% { 'A1', 'B1', ...
%   'A2', 'B2', ...
%   'A3', .... }
function refs = create_cellref_table(size_celldata)

size_celldata = size_celldata(1 : 2);

colstr = colind_to_name(1 : size_celldata(2));
colstr = repmat(colstr, size_celldata(1), 1);
rowstr = cellfun(@num2str, num2cell((1 : size_celldata(1))'), 'UniformOutput', false);
rowstr = repmat(rowstr, 1, size_celldata(2));

refs = strcat(colstr, rowstr);

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


%%
function mkdir_p(varargin)

narginchk(1, inf);

if nargin == 1 && exist(varargin{1}, 'dir')
    return
elseif nargin >= 2 && exist(fullfile(varargin{:}), 'dir')
    return
end

mkdir(fullfile(varargin{:}));

end


%%
function try_movefile(src, dest)

for trial = 1 : 3
    [status, msg, msgid] = movefile(src, dest, 'f');
    % Unix コマンドと違って status は成功時に 1
    % exist の file はフォルダも含む
    if status && exist(dest, 'file')
        return
    elseif ~ isempty(strfind(msgid, 'FileDoesNotExist'))
        throw(msg);
    end
    pause(0.2);
end

error(msg);

end


%%
function try_rmdir(dirname)
% try_rmdir(dirname)

if ~ exist(dirname, 'dir')
    return
end
    
for trial = 1 : 3
    [status, msg, ~] = rmdir(dirname, 's');
    % Unix コマンドと違って成功時に 1
    if status
        return
    end
    pause(0.2);
end

% 削除できなくても問題ないケースを想定しているので
% エラーではなくワーニングで
warning(msg);

end
