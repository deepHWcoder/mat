function exportdoc()
% �趨����Word�ļ�����·��
cd
filespec_user = [pwd '\�Զ�����.doc'];

% �ж�Word�Ƿ��Ѿ��򿪣����Ѵ򿪣����ڴ򿪵�Word�н��в���������ʹ�Word
try
    % ��Word�������Ѿ��򿪣���������Word
    Word = actxGetRunningServer('Word.Application');
catch
    % ����һ��Microsoft Word�����������ؾ��Word
    Word = actxserver('Word.Application'); 
end;

% ����Word����Ϊ�ɼ�
Word.Visible = 1;    % ��set(Word, 'Visible', 1);

% �������ļ����ڣ��򿪸ò����ļ��������½�һ���ļ��������棬�ļ���Ϊ����.doc
if exist(filespec_user,'file'); 
    Document = Word.Documents.Open(filespec_user);     
    % Document = invoke(Word.Documents,'Open',filespec_user);
else
    Document = Word.Documents.Add;     
    % Document = invoke(Word.Documents, 'Add'); 
    Document.SaveAs(filespec_user);
end

Content = Document.Content;    % ����Content�ӿھ��
Selection = Word.Selection;    % ����Selection�ӿھ��
Paragraphformat = Selection.ParagraphFormat;  % ����ParagraphFormat�ӿھ��

% ҳ������
Document.PageSetup.TopMargin = 60;      % �ϱ߾�60��
Document.PageSetup.BottomMargin = 45;   % �±߾�45��
Document.PageSetup.LeftMargin = 45;     % ��߾�45��
Document.PageSetup.RightMargin = 45;    % �ұ߾�45��

% �ڹ������λ�ò���һ��12��9�еı��  
Tables = Document.Tables.Add(Selection.Range,17,1); 
% ���ص�1�����ľ��  
DTI = Document.Tables.Item(1);    % ��DTI = Tables;   
% ���ñ��߿�
DTI.Borders.OutsideLineStyle = 'wdLineStyleSingle'; 
DTI.Borders.OutsideLineWidth = 'wdLineWidth150pt'; 
DTI.Borders.InsideLineStyle = 'wdLineStyleSingle'; 
% DTI.Borders.InsideLineWidth = 'wdLineWidth150pt'; 
DTI.Rows.Alignment = 'wdAlignRowCenter'; 
% DTI.Rows.Item(8).Borders.Item(1).LineStyle = 'wdLineStyleNone'; 
% DTI.Rows.Item(8).Borders.Item(3).LineStyle = 'wdLineStyleNone'; 
% DTI.Rows.Item(11).Borders.Item(1).LineStyle = 'wdLineStyleNone'; 
% DTI.Rows.Item(11).Borders.Item(3).LineStyle = 'wdLineStyleNone';

% ���ñ���п���и�  
% column_width = [53.7736,85.1434,53.7736,35.0094];    
% �����п����� 
% row_height = [28.5849,28.5849,28.5849,28.5849,25.4717,25.4717,32.8302,312.1698,17.8302,49.2453,14.1509,18.6792]; 
% % �����и����� % ͨ��ѭ�����ñ��ÿ�е��п� 
% for i = 1:9
%     DTI.Columns.Item(i).Width = column_width(i);
% end
% ͨ��ѭ�����ñ��ÿ�е��и� 
% for i=1:17
%     DTI.Rows.Item(i).Height = row_height(i);
%     DTI.Cell(i,j).VerticalAlignment = 'wdCellAlignVerticalCenter'; 
% end





% �趨�ĵ����ݵ���ʼλ�úͱ���
str2 = 'AAA';
Content.Start = 0;         % �����ĵ����ݵ���ʼλ��
title = str2;
Content.Text = title;      % ������������
Content.Font.Size = 16 ;   % �����ֺ�Ϊ16
Content.Font.Bold = 4 ;    % ����Ӵ�
Content.Paragraphs.Alignment = 'wdAlignParagraphCenter';    % ���ж���

Selection.Start = Content.end;    % �趨�������ݵ���ʼλ��
Selection.TypeParagraph;    % �س�������һ��

aa = strcat('  ',str2);
Selection.Text = aa;     % �ڵ�ǰλ��������������
Selection.Font.Size = 12;   % �����ֺ�Ϊ12
Selection.Font.Bold = 0;    % ���岻�Ӵ�
Selection.MoveDown;         % ������ƣ�ȡ��ѡ�У�
paragraphformat.Alignment = 'wdAlignParagraphCenter';    % ���ж���
Selection.TypeParagraph;    % �س�������һ��
Selection.TypeParagraph;    % �س�������һ��
Selection.Font.Size = 10.5; % �����ֺ�Ϊ10.5

Selection.Start = Content.end;    % �趨�������ݵ���ʼλ��
Selection.PasteSpecial;
% delete(zft);    % ɾ��ͼ�ξ��

 

Document.ActiveWindow.ActivePane.View.Type = 'wdPrintView';    % ������ͼ��ʽΪҳ��
Document.Save;     % �����ĵ�
end
xlswrite('AAAA.xls',OutputData(:,(1:n))','sheet1',['C3:',setstr('C'+4),num2str(3+n-1)]);