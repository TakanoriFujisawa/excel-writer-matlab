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
