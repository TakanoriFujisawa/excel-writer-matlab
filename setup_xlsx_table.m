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


