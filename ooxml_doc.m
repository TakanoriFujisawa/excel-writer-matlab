
% Office 2007 BIFF12 Œ`Ž®
% web('http://www.arstdesign.com/articles/office2007bin.html', '-browser')

if ispc
    winopen('../../study/misc/xlsx/ooxml.pdf');
elseif ismac
    system('open ../../study/misc/xlsx/ooxml.pdf');
else
    system('xdg-open ../../study/misc/xlsx/ooxml.pdf');
end
