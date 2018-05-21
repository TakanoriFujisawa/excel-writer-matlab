
hfig = figure(1);

n_children = length(hfig.Children); 
for i = 1 : length(hfig.Children)
    hfig.Children(1).delete
end
clf reset;

%uitab('Title', '<html><b>Html</b></html>')

%[jlabel, hlabel] = javacomponent('javax.swing.JLabel')
[jpane, hpane] = javacomponent('javax.swing.JScrollPane');
hpane.Position = [ 5, 0, hfig.Position(3) - 5, hfig.Position(4) ];
hfig.SizeChangedFcn = @(hfig, ~) set(hpane, 'Position', [5, 0, hfig.Position(3) - 5, hfig.Position(4) ]);


jpane.setHorizontalScrollBarPolicy(javax.swing.JScrollPane.HORIZONTAL_SCROLLBAR_AS_NEEDED);

%[jtable, htable] = javacomponent('javax.swing.JTable');
jtable = javax.swing.JTable(40, 50);
jtable.setAutoResizeMode(javax.swing.JTable.AUTO_RESIZE_OFF);
jtable.setColumnSelectionAllowed(true);
jtable.setRowSelectionAllowed(true);

%htable.Position(1 : 2) = [20 20];
%htable.Position(3 : 4) = hfig.Position(3 : 4) - 40;

%[jheader, hheader] = javacomponent(jtable.getTableHeader);
%hheader.Position = [20, 20 + htable.Position(4), htable.Position(3), 20];


model = jtable.getModel();


%{
row = javaArray('java.lang.String', 3);
for i = 1 : 3
    row(i) = java.lang.String(num2str(i));
end
%}
row = {'hoge', 'fuga', 'moga', 'matsu'};

for i = 1 : 5
    model.addRow(row);
end

jpane.setViewportView(jtable);
