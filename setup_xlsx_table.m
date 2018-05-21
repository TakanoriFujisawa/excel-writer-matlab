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


