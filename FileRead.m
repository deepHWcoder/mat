cd(pathname);
fileName       = dir(strcat(pathname,'*.mat')); 
n                  = length(fileName);                                
for i=1:n
fileNameTemp   = fileName(i,1).name;  
[fidin,message]    = fopen([pathname,fileNameTemp],'r');     % 打开文件  
    if fidin    == -1
        error('AAAA');
        else        
            while ~feof(fidin)                                     % 判断是否为文件末尾              
                 tline=fgetl(fidin);                                % 从文件读行  
                if length(tline)>5                %只判断有效的字符与数据行，例如认为5个字符以下的行都为与数据无关行，可以忽略                    
                        Pos_space   = strfind(tline,' ');            %以空格作为区分两变量的标志：找出空格的位置，则空格位置之间的长度就是变量名称
                        num_VarName = length(Pos_space);
                        
                       for j=1:num_VarName-1                            %寻找所需变量的位置
                           if strcmp(tline(Pos_space(j)+1:Pos_space(j+1)-1),'time')==1 
                               Pos_time     = j+1;continue;                               
                           end                          
                        end                                          
                 end
            end 
    end
fclose(fidin);