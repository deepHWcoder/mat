filereport = [cd, '\FR.xls'];

         try
             Excel = actxGetRunningServer('Excel.Application');
         catch
             Excel = actxserver('Excel.Application'); 
         end

         set(Excel, 'Visible', 1);  

         if exist(filereport,'file'); 
            Workbook = invoke(Excel.Workbooks,'Open',filereport);
         else
            Workbook = invoke(Excel.Workbooks, 'Add');
            Workbook.SaveAs(filereport);
         end

         Sheets = Excel.ActiveWorkbook.Sheets;   
         Sheet1 = Sheets.Item(1); 
         Sheet1.Activate;
          
         Activesheet=Excel.ActiveSheet;
         ActivesheetRange = get(Activesheet,'Range','A1:L50');
         A = [];
         set(ActivesheetRange, 'Value', A);
          
         Shapes=Excel.ActiveSheet.Shapes;
         if Shapes.Count~=0;
            for i=1:Shapes.Count;
                Shapes.Item(1).Delete;
            end;
         end;


         Sheet1.PageSetup.TopMargin = 60; 
         Sheet1.PageSetup.BottomMargin = 45; 
         Sheet1.PageSetup.LeftMargin = 45;
         Sheet1.PageSetup.RightMargin = 45;

 
         Sheet1.Range('A1:A26').RowHeight = 18;
         Sheet1.Range('A1:J1').ColumnWidth = [70,13,10,10,10,10,10,10,10,10];


         Sheet1.Range('A1').Font.size = 13;
         Sheet1.Range('A1').Font.bold = 2;

         Sheet1.Range('A1').Value                                = {'Swept Steer£º'};

      
     
         i=1;
         if i<= N
             
     
             AVERAGE_TEST_SPEED(i)                               = roundn(AVERAGE_TEST_SPEED(i),-4);
             Sheet1.Range('C3').Value                            = {AVERAGE_TEST_SPEED(i)};
             Sheet1.Range('B3').Value                            = roundn(Sheet1.Range('C3').Value,-4);
                           
             figureA(i) = figure('units','normalized','position',[0.1 0.1 0.4 0.6],'visible','off');
             subplot(311);                 
                 subplot(312);                 
                 subplot(313);                            
           
             hgexport(figureA(i), '-clipboard');
             Excel.ActiveSheet.Range('K2').Select;                              
             Excel.ActiveSheet.Paste;
             delete(figureA(i));  
             
           
             
             figureA(i+10) = figure('units','normalized','position',[0.1 0.1 0.4 0.6],'visible','off');
             subplot(311);                
                 subplot(312);                 
                 subplot(313);                 
             hgexport(figureA(i+10),'-clipboard');
             Excel.ActiveSheet.Range('K28').Select;                              
             Excel.ActiveSheet.Paste;
             delete(figureA(i+10));               
             
             i=i+1;

             
         end   
       
         Workbook.Save;  
         delete(Excel); 