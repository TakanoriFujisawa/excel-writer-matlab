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
% �f�[�^�ǉ�
%  1. �V�[�g�̐��𒲂ׂ�
%  2. �^�u������Ȃ��ꍇ�̓^�u�E�e�[�u����ǉ�
%  3. �^�u��Title�ɃV�[�g�����Z�b�g����
%  4. �e�[�u���� CellEditCallback ���Z�b�g
%  5. �^�u�Ƀf�[�^��ǉ� (�f�[�^�𕪔z)
%
% �^�u�Ƀf�[�^��ǉ�
%  1. �K�v����ColumnName�𐶐����ēo�^
%  2. �s�������� { 'char' } �Z���z��𐶐����� ColumnFormat �ɃZ�b�g
% [3. �f�[�^�𕶎���ɕϊ� (�ҏW���\�ɂ���ꍇ)]
%  4. �f�[�^���Œ� 20 �s 6 ��ɂȂ�悤�g�� (���h���̂���)
%  5. �f�[�^���Z�b�g
%

if nargout == 0
    warning('[Internal] �X�V��̃n���h����Ԃ�l�Ŏ󂯎��K�v������܂�');
end

if ishandle(hdl) && strcmp(hdl.Visible, 'off')
    return
elseif ~ ishandle(hdl)
    % Figure �������Ă��鎞
    hdl = setup_xlsx_table([], xlsx);
end


for t = 1 : size(xlsx.data, 3)
    if t > length(hdl.Children.Children)
        % ����Ȃ��^�u��ǉ�
        htab = uitab('Parent', hdl.Children, 'Title', ['Sheet', num2str(t)]);
        htbl = uitable('Parent', htab, 'Units', 'normalized', 'Position', [0 0 1 1], 'RowName', 'numbered');
    else
        htab = hdl.Children.Children(t);
        htbl = htab.Children(1);
    end
    
    htab.Title = xlsx.get_sheet_name_from_index(t);
    
    % �\���p�̃f�[�^�̎擾
    data = xlsx.data(:, :, t);
    
    % ���h���̂��߃T�C�Y�� 20x6 �܂Ŋg��
    if size(data, 1) < 20
        data(end + 1 : 20, :) = {''};
    end
    
    if size(data, 2) < 6
        data(:, end + 1 : 6) = {''};
    end
    
    % ColumnFormat �� ColumnName ��K�v���p��
    if ~ iscell(htbl.ColumnFormat) || length(htbl.ColumnFormat) < size(data, 2)
        htbl.ColumnFormat = repmat({'char'}, 1, size(data, 2) * 2);
    end
    
    if ~ iscell(htbl.ColumnName) || length(htbl.ColumnName) < size(data, 2)
        htbl.ColumnName = colind_to_name(1 : size(data, 2) + 3);
    end
    
    % �f�[�^���Z�b�g
    htbl.Data = data;
end

end
