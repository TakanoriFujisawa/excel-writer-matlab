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

