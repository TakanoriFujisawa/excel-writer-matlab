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
