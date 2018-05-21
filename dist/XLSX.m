classdef XLSX < handle
    properties
        data = {}
        sheet_name = {}
        show_table = false
        hdl_table = []
    end
    
    methods
        % �I�[�o�[���[�h���ׂ����\�b�h
        %
        % subsref, subasgn
        % end
        % numel, size
        % double, char (�^�ϊ�)
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
            % ind �Ԗڂ̃V�[�g�̖��O��Ԃ��C�ݒ肳��Ă��Ȃ����
            % �f�t�H���g�̖��O��Ԃ�
            %
            %  % �X�J���[���͎�
            %  xl.get_sheet_name_from_index(1)     => 'Sheet1'
            %
            %  % �x�N�g�����͎�
            %  xl.get_sheet_name_from_index(2 : 3) => {'Sheet2', 'Sheet3'}
            validateattributes(ind, {'numeric'}, {'positive', 'integer'});
            
            if any(ind > size(this.data, 3))
                error('�C���f�b�N�X���V�[�g�͈̔͊O�ł�');
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
                error('�V�[�g���͕�����܂��͕�����̃Z���z�񂪕K�v�ł�');
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
            % subs : subsasgn �� S (struct)
            % { [1 : 2], ':', 2 }
            % { [3x3 logical] }
            %
            % ����� :
            % x(1, 2 : 3) = 'str'
            % x(2, 4) = 1;
            % x(2, 4 : 6) = 1 : 3;
            %
            %
            % x(3, 2) = rand(5);
            % x(4, 2) = {'str', 2, true}
            %  [���̏����̏ꍇ�C���f�b�N�X�̓X�J���[�łȂ���΂����Ȃ�]
            %
            
            if isempty(val)
                val = { val };
                
            elseif ischar(val)
                val = { val(:)' };
                
            elseif isnumeric(val) || islogical(val)
                if ndims(val) > 3
                    error('���͍͂ő�R�����̔z��ł���K�v������܂�');
                end
                val = num2cell(val);
                
            elseif iscell(val)
                if ndims(val) > 3
                    error('���͍͂ő�R�����̔z��ł���K�v������܂�');
                end
                
                if ~ all(cellfun(@isxlsxdata, val))
                    error('�Z���̗v�f�͐��l�C�_���l�C�܂��͕�����ł���K�v������܂�');
                end
                
            else
                error('���͕͂�����E���l�z��܂��̓Z���z�񂪕K�v�ł�');
            end
            
            if nargin == 3 && ~ isempty(subs)
                % cell �ɕϊ����� subsasgn �ɓn��
                %this.data = subsasgn(this.data, subs, val);
                switch subs.type
                    case {'()', '{}'}
                        switch length(subs.subs)
                            case {1, 2}
                                % �Y������̃��[�������G�����Ă悭������Ȃ������̂�
                                % ��U 1 �����X���C�X�ɑ΂��� subsasgn �������㍇��
                                sz = [size(this.data, 1), size(this.data, 2), size(this.data, 3)];
                                tmp_data = subsasgn(this.data(:, :, 1), subs, val);
                                tmp_data(1 : sz(1), 1 : sz(2), 2 : sz(3)) = this.data(:, :, 2 : end);
                                
                            case 3
                                tmp_data = subsasgn(this.data, subs, val);
                                
                            otherwise
                                error(message('MATLAB:badsubscriptTextDimension'));
                        end
                        
                        % 4 �����z��ȏ�ɂȂ�Ȃ��悤��U tmp_data �ɑ������
                        % �G���[�`�F�b�N
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
            
            % �t�@�C������ (�f�[�^�̑����ɒ��ӁI)
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
                    error('XLSX �`���ȊO�̓ǂݎ��̓T�|�[�g���Ă��܂���');
                end
                
                [this.data, this.sheet_name] = read_xlsx(varargin{1});
                return
            end
            
            if isempty(varargin{1})
                % do nothing
                
            elseif isnumeric(varargin{1}) || islogical(varargin{1}) || iscell(varargin{1})
                this.set_data(varargin{1});
                
            else
                error('��P�����͐��l�z��E�_���l�z��C�܂��̓Z���z����w�肵�܂�');
                
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
                            error('�s���ȃv���p�e�B %s �ւ̎Q�Ƃł�', subs(1).subs);
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
                    % �f�[�^�����̃V�[�g�ɓ������Ⴄ���Ƃ�����̂ł��܂萄�����Ȃ�
                    ind = numel(this.data);
                    
                case 2
                    % x(1, end) = 3
                    % 3�����z��̏ꍇ�ƈ���đ�1�Z���̃C���f�b�N�X��Ԃ�
                    ind = size(this.data, the_indexed);
                case 3
                    % x(1:end, 1:end, 1) = 1
                    ind = size(this.data, the_indexed);
                    
                otherwise
                    error('�Y���̐����������܂�')
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
            
            % �g���q��ǉ�
            [directory, outfile, ext] = fileparts(outfile);
            if isempty(ext) || ~ strcmpi(ext, '.xlsx')
                outfile = fullfile(directory, [outfile, ext, '.xlsx']);
            else
                outfile = fullfile(directory, [outfile, ext]);
            end
            
            %% �o�̓Z���̃`�F�b�N
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
% ��ԍ����A���t�@�x�b�g�\�L�ɕϊ� 1 -> A
% ���͂̓X�J���[���R���̂�

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
    % [A=0 ~ Z=25 �܂ł� 26 �i�� 2 ��] + 27
    ind = ind - 27;
    name = char('A' + [ floor(ind / 26), mod(ind, 26) ]);
    
elseif ind < 17575  % �ꉞ 16384(XFD) �܂ł炵��
    % [A=0 ~ Z=25 �܂ł� 26 �i�� 3 ��] + 703
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


%%
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
            % xlswrite �̎d�l�ɍ��킹�� 65535 (-65535) �ɕϊ�
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
% �Q�Ɨp�� A1 �X�^�C���̈ȉ��̕\�����쐬 (Cell reference)
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
    % Unix �R�}���h�ƈ���� status �͐������� 1
    % exist �� file �̓t�H���_���܂�
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
    % Unix �R�}���h�ƈ���Đ������� 1
    if status
        return
    end
    pause(0.2);
end

% �폜�ł��Ȃ��Ă����Ȃ��P�[�X��z�肵�Ă���̂�
% �G���[�ł͂Ȃ����[�j���O��
warning(msg);

end
