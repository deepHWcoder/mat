%用Matlab生成Word文档
%用Matlab编了一段程序，可以生成Word文档，文档中含有表格，代码如下：
filespec = ['自动报告' datestr(now,30) '.doc'];
try;
    Word=actxGetRunningServer('Word.Application');
catch;
    Word = actxserver('Word.Application'); 
end;
set(Word, 'Visible', 1);
documents = Word.Documents;
if exist(filespec,'file')
    document = invoke(documents,'Open',filespec);    
else
    document = invoke(documents, 'Add');
    document.SaveAs(filespec);
end
content = document.Content;
duplicate = content.Duplicate;
inlineshapes = content.InlineShapes;
selection = Word.Selection;
paragraphformat = selection.ParagraphFormat;
%页面设置
document.PageSetup.TopMargin = 60;
document.PageSetup.BottomMargin = 45;
document.PageSetup.LeftMargin = 40;
document.PageSetup.RightMargin = 40;
set(content, 'Start',0);
title='自动化报告';
set(content, 'Text',title);
set(paragraphformat, 'Alignment','wdAlignParagraphCenter');
rr=document.Range(0,16);%选择文本
rr.Font.Size=20;%设置文本字体
rr.Font.Bold=4;%设置文本字体
end_of_doc = get(content,'end');
set(selection,'Start',end_of_doc);
selection.Font.Size=10;
selection.MoveDown;
selection.TypeParagraph;
selection.TypeParagraph;
% set(paragraphformat, 'Alignment','wdAlignParagraphCenter');
selection.Font.Size=10.5;
Tables=document.Tables.Add(selection.Range,7,12);

%设置边框
DTI=document.Tables.Item(1);
DTI.Borders.OutsideLineStyle='wdLineStyleSingle';
DTI.Borders.OutsideLineWidth='wdLineWidth150pt';
DTI.Borders.InsideLineStyle='wdLineStyleSingle';
DTI.Borders.InsideLineWidth='wdLineWidth150pt';
DTI.Rows.Alignment='wdAlignRowCenter';
end_of_doc = get(content,'end');
set(selection,'Start',end_of_doc);
selection.TypeParagraph;
set(selection, 'Text','       年    月    日');
set(paragraphformat, 'Alignment','wdAlignParagraphRight');
end_of_doc = get(content,'end');
set(selection,'Start',end_of_doc);
DTI.Rows.Item(5).Borders.Item(1).LineStyle='wdLineStyleNone';
column_width=[60, 45, 45, 45, 45, 45, 45, 45, 45, 45, 45, 45];
row_height=[20, 30, 30, 40, 400, 40, 40];
for i = 1 : 12
    DTI.Columns.Item(i).Width =column_width(i);
end
for i=1:7
    DTI.Rows.Item(i).Height =row_height(i);
end
for i = 1 : 7
    for j = 1 : 12
        DTI.Cell(i, j).VerticalAlignment='wdCellAlignVerticalCenter';
    end
end
DTI.Cell(1, 2).Merge(DTI.Cell(1, 12));
DTI.Cell(4, 1).Merge(DTI.Cell(4, 12));
DTI.Cell(5, 1).Merge(DTI.Cell(5, 12));
DTI.Cell(6, 1).Merge(DTI.Cell(6, 12));
DTI.Cell(7, 1).Merge(DTI.Cell(7, 12));
DTI.Cell(1, 1).Range.Text = '测试时间';
DTI.Cell(1, 2).Range.Text = datestr(now, 31);
DTI.Cell(2, 1).Range.Text = '所有项目';
DTI.Cell(4,1).Range.Text = '自动化测试报告:';
DTI.Cell(4,1).Range.ParagraphFormat.Alignment='wdAlignParagraphLeft';
DTI.Cell(5, 1).Range.ParagraphFormat.Alignment='wdAlignParagraphLeft';
DTI.Cell(5, 1).VerticalAlignment='wdCellAlignVerticalTop';
DTI.Cell(6, 1).Range.Text = '   测试员签字 :                        年    月    日';

