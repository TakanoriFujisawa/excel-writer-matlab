function mkdir_p(varargin)

narginchk(1, inf);

if nargin == 1 && exist(varargin{1}, 'dir')
    return
elseif nargin >= 2 && exist(fullfile(varargin{:}), 'dir')
    return
end

mkdir(fullfile(varargin{:}));

end
