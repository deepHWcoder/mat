cd(pathname);
fileName       = dir(strcat(pathname,'*.mat')); 
n                  = length(fileName);                                
for i=1:n
fileNameTemp   = fileName(i,1).name;  
[fidin,message]    = fopen([pathname,fileNameTemp],'r');     % ���ļ�  
    if fidin    == -1
        error('AAAA');
        else        
            while ~feof(fidin)                                     % �ж��Ƿ�Ϊ�ļ�ĩβ              
                 tline=fgetl(fidin);                                % ���ļ�����  
                if length(tline)>5                %ֻ�ж���Ч���ַ��������У�������Ϊ5���ַ����µ��ж�Ϊ�������޹��У����Ժ���                    
                        Pos_space   = strfind(tline,' ');            %�Կո���Ϊ�����������ı�־���ҳ��ո��λ�ã���ո�λ��֮��ĳ��Ⱦ��Ǳ�������
                        num_VarName = length(Pos_space);
                        
                       for j=1:num_VarName-1                            %Ѱ�����������λ��
                           if strcmp(tline(Pos_space(j)+1:Pos_space(j+1)-1),'time')==1 
                               Pos_time     = j+1;continue;                               
                           end                          
                        end                                          
                 end
            end 
    end
fclose(fidin);