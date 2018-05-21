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
            %
            % <a href="matlab:doc('XLSX')">XLSX</a>
            
            % ファイルから (データの損失に注意！)
            % x = XLSX(xls_filename);
            
            % "for internal use only"
            % x = XLSX(data, sheet_name, 'valid')
            
            
            if nargin == 0
                return
            end
            
            if nargin == 3 && strcmp(varargin{3}, 'valid')
                this.data = varargin{1};
                this.sheet_name = varargin{3};
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
