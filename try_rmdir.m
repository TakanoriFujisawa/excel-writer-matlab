function try_rmdir(dirname)
% try_rmdir(dirname)

if ~ exist(dirname, 'dir')
    return
end
    
for trial = 1 : 3
    [status, msg, ~] = rmdir(dirname, 's');
    % Unix コマンドと違って成功時に 1
    if status
        return
    end
    pause(0.2);
end

% 削除できなくても問題ないケースを想定しているので
% エラーではなくワーニングで
warning(msg);

end
