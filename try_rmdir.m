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
