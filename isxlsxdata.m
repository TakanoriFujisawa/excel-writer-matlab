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

