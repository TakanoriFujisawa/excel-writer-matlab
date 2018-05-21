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
