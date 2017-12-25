# utl_excel_import_two_excel_ranges_within_one_sheet
Excel import two excel ranges within one sheet.                                                                                                           Python/SAS/WPS: Worked fine on my $600 desktop importing a tiny 140mb csv file                                                                             Keywords: sas sql join merge big data analytics macros oracle teradata mysql sas communities stackoverflow statistics                                    artificial inteligence AI Python R Java Javascript WPS Matlab SPSS Scala Perl C C# Excel MS Access JSON graphics maps NLP                                 natural language processing machine learning.    

    Excel import two excel ranges within one sheet;                                               
                                                                                                  
    related to                                                                                    
    Original post                                                                                 
    Switch from DDE to something that works, probably like libname                                
                                                                                                  
    INPUT                                                                                         
    =====                                                                                         
                                                                                                  
      d:/xls/class.xlsx                                                                           
                                                                                                  
          <       RANGE A1:D3                                 >    <     RANGE F1-G3        >     
          +--------------------------------------------------------+--------------------------    
          |     A      |    B       |     C      |    D       | _  |     F      |    G       |    
          +--------------------------------------------------------+--------------------------    
       1  | NAME       |   SEX      |    AGE     |  HEIGHT    |    | NAME1      |   SEX1     |    
          +------------+------------+------------+------------+----+------------+------------+    
       2  | ALFRED     |    M       |    15      |    69      |    | ALFRED     |    M       |    
          +------------+------------+------------+------------+----+------------+------------+    
       3  | ALICE      |    F       |    13      |    58      |    | ALICE      |    F       |    
          +------------+------------+------------+------------+----+------------+------------+    
                                                                                                  
    [CLASS]                                                                                       
                                                                                                  
                                                                                                  
    PROCESS                                                                                       
    =======                                                                                       
                                                                                                  
       %symdel range / nowarn;                                                                    
       libname xel "d:/xls/class.xlsx";                                                           
       data _null_;                                                                               
                                                                                                  
         do range= 'A1:D3','F1:G3';                                                               
           call symputx('range',range);                                                           
           rc=dosubl('                                                                            
               data %substr(&range.,1,2);                                                         
                  set xel."class$&range."n;                                                       
               run;quit;                                                                          
           ');                                                                                    
         end;                                                                                     
                                                                                                  
       run;quit;                                                                                  
       libname xel clear;                                                                         
                                                                                                  
                                                                                                  
    OUTPUT(Two datasets onr for each range)                                                       
    =========================================                                                     
                                                                                                  
      WORK.A1 total obs=2                                                                         
                                                                                                  
         NAME     SEX    AGE    HEIGHT                                                            
                                                                                                  
        Alfred     M      14     69.0                                                             
        Alice      F      13     56.5                                                             
                                                                                                  
                                                                                                  
      WORK.F1 total obs=2                                                                         
                                                                                                  
         NAME     SEX    AGE                                                                      
                                                                                                  
        Alfred     M      14                                                                      
        Alice      F      13                                                                      
                                                                                                  
    *                _               _       _                                                    
     _ __ ___   __ _| | _____     __| | __ _| |_ __ _                                             
    | '_ ` _ \ / _` | |/ / _ \   / _` |/ _` | __/ _` |                                            
    | | | | | | (_| |   <  __/  | (_| | (_| | || (_| |                                            
    |_| |_| |_|\__,_|_|\_\___|   \__,_|\__,_|\__\__,_|                                            
                                                                                                  
    ;                                                                                             
                                                                                                  
    %utlfkil(d:/xls/class.xlsx);                                                                  
    libname xel "d:/xls/class.xlsx";                                                              
    data xel.class;                                                                               
      set sashelp.class(obs=2);                                                                   
         rename weight=_;                                                                         
         name1=name;                                                                              
         sex1=sex;                                                                                
         weight=.;                                                                                
    run;quit;                                                                                     
    libname xel clear;                                                                            
                                                                                                  
    *          _       _   _                                                                      
     ___  ___ | |_   _| |_(_) ___  _ __                                                           
    / __|/ _ \| | | | | __| |/ _ \| '_ \                                                          
    \__ \ (_) | | |_| | |_| | (_) | | | |                                                         
    |___/\___/|_|\__,_|\__|_|\___/|_| |_|                                                         
                                                                                                  
    ;                                                                                             
                                                                                                  
    %symdel range / nowarn;                                                                       
    libname xel "d:/xls/class.xlsx";                                                              
    data _null_;                                                                                  
                                                                                                  
      do range= 'A1:D3','F1:G3';                                                                  
        call symputx('range',range);                                                              
        rc=dosubl('                                                                               
            data %substr(&range.,1,2);                                                            
               set xel."class$&range."n;                                                          
            run;quit;                                                                             
        ');                                                                                       
      end;                                                                                        
                                                                                                  
    run;quit;                                                                                     
    libname xel clear;                                                                            
                                                                                                  
                                                                                                  
    SYMBOLGEN:  Macro variable RANGE resolves to A1:D3                                            
    SYMBOLGEN:  Macro variable RANGE resolves to A1:D3                                            
    NOTE: There were 2 observations read from the data set XEL.'class$A1:D3'n.                    
    NOTE: The data set WORK.A1 has 2 observations and 4 variables.                                
    NOTE: DATA statement used (Total process time):                                               
          real time           0.01 seconds                                                        
                                                                                                  
    SYMBOLGEN:  Macro variable RANGE resolves to F1:G3                                            
    SYMBOLGEN:  Macro variable RANGE resolves to F1:G3                                            
    NOTE: There were 2 observations read from the data set XEL.'class$F1:G3'n.                    
    NOTE: The data set WORK.F1 has 2 observations and 2 variables.                                
    NOTE: DATA statement used (Total process time):                                               
          real time           0.00 seconds                                                        
