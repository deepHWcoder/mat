function exportdoc()
% 设定测试Word文件名和路径
cd
filespec_user = [pwd '\自动报告.doc'];

% 判断Word是否已经打开，若已打开，就在打开的Word中进行操作，否则就打开Word
try
    % 若Word服务器已经打开，返回其句柄Word
    Word = actxGetRunningServer('Word.Application');
catch
    % 创建一个Microsoft Word服务器，返回句柄Word
    Word = actxserver('Word.Application'); 
end;

% 设置Word属性为可见
Word.Visible = 1;    % 或set(Word, 'Visible', 1);

% 若测试文件存在，打开该测试文件，否则，新建一个文件，并保存，文件名为测试.doc
if exist(filespec_user,'file'); 
    Document = Word.Documents.Open(filespec_user);     
    % Document = invoke(Word.Documents,'Open',filespec_user);
else
    Document = Word.Documents.Add;     
    % Document = invoke(Word.Documents, 'Add'); 
    Document.SaveAs(filespec_user);
end

Content = Document.Content;    % 返回Content接口句柄
Selection = Word.Selection;    % 返回Selection接口句柄
Paragraphformat = Selection.ParagraphFormat;  % 返回ParagraphFormat接口句柄

% 页面设置
Document.PageSetup.TopMargin = 60;      % 上边距60磅
Document.PageSetup.BottomMargin = 45;   % 下边距45磅
Document.PageSetup.LeftMargin = 45;     % 左边距45磅
Document.PageSetup.RightMargin = 45;    % 右边距45磅

% 在光标所在位置插入一个12行9列的表格  
Tables = Document.Tables.Add(Selection.Range,17,1); 
% 返回第1个表格的句柄  
DTI = Document.Tables.Item(1);    % 或DTI = Tables;   
% 设置表格边框
DTI.Borders.OutsideLineStyle = 'wdLineStyleSingle'; 
DTI.Borders.OutsideLineWidth = 'wdLineWidth150pt'; 
DTI.Borders.InsideLineStyle = 'wdLineStyleSingle'; 
% DTI.Borders.InsideLineWidth = 'wdLineWidth150pt'; 
DTI.Rows.Alignment = 'wdAlignRowCenter'; 
% DTI.Rows.Item(8).Borders.Item(1).LineStyle = 'wdLineStyleNone'; 
% DTI.Rows.Item(8).Borders.Item(3).LineStyle = 'wdLineStyleNone'; 
% DTI.Rows.Item(11).Borders.Item(1).LineStyle = 'wdLineStyleNone'; 
% DTI.Rows.Item(11).Borders.Item(3).LineStyle = 'wdLineStyleNone';

% 设置表格列宽和行高  
% column_width = [53.7736,85.1434,53.7736,35.0094];    
% 定义列宽向量 
% row_height = [28.5849,28.5849,28.5849,28.5849,25.4717,25.4717,32.8302,312.1698,17.8302,49.2453,14.1509,18.6792]; 
% % 定义行高向量 % 通过循环设置表格每列的列宽 
% for i = 1:9
%     DTI.Columns.Item(i).Width = column_width(i);
% end
% 通过循环设置表格每行的行高 
% for i=1:17
%     DTI.Rows.Item(i).Height = row_height(i);
%     DTI.Cell(i,j).VerticalAlignment = 'wdCellAlignVerticalCenter'; 
% end





% 设定文档内容的起始位置和标题
str2 = 'AAA';
Content.Start = 0;         % 设置文档内容的起始位置
title = str2;
Content.Text = title;      % 输入文字内容
Content.Font.Size = 16 ;   % 设置字号为16
Content.Font.Bold = 4 ;    % 字体加粗
Content.Paragraphs.Alignment = 'wdAlignParagraphCenter';    % 居中对齐

Selection.Start = Content.end;    % 设定下面内容的起始位置
Selection.TypeParagraph;    % 回车，另起一段

aa = strcat('  ',str2);
Selection.Text = aa;     % 在当前位置输入文字内容
Selection.Font.Size = 12;   % 设置字号为12
Selection.Font.Bold = 0;    % 字体不加粗
Selection.MoveDown;         % 光标下移（取消选中）
paragraphformat.Alignment = 'wdAlignParagraphCenter';    % 居中对齐
Selection.TypeParagraph;    % 回车，另起一段
Selection.TypeParagraph;    % 回车，另起一段
Selection.Font.Size = 10.5; % 设置字号为10.5

Selection.Start = Content.end;    % 设定下面内容的起始位置
Selection.PasteSpecial;
% delete(zft);    % 删除图形句柄

 

Document.ActiveWindow.ActivePane.View.Type = 'wdPrintView';    % 设置视图方式为页面
Document.Save;     % 保存文档
end
xlswrite('AAAA.xls',OutputData(:,(1:n))','sheet1',['C3:',setstr('C'+4),num2str(3+n-1)]);