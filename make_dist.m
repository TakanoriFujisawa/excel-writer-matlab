
files = {
    'XLSX.m';
    'colind_to_name.m';
    'isxlsxdata.m';
    'setup_xlsx_table.m'
    'update_xlsx_table.m'
    'write_xlsx.m'
    'mkdir_p.m'
    'try_movefile.m'
    'try_rmdir.m'
    };


fid = fopen('dist/XLSX.m', 'w');

for f = 1 : length(files)
    if f > 1
        fprintf(fid, '\n\n%%%%\n');
    end
    
    text = fileread(files{f});
    
    fprintf(fid, '%s', text);
end

fclose(fid);
