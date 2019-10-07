'******************************************************************************
'* File             :pdm2xls.vbs 
'* version          :v17 dd 20141114
'* File Type        :VB Script File (needs Powerdesigner plus excel environment installed) exec {CTRL+SHIFT+x}
'* Purpose          :Export  Logical Model to Excel from PDM model     
'* Title            :pdm2xls
'* Category         :BI&D ontwerp script
'* Identificatie    :\BI&D_EDW_MTHV\1130 Hulpmiddelen\3000 Scripts\3200 Ontwerp 

'Big thanks to cc.wust@belastingdienst.nl  
'Hello to nj.essenstam@belastingdienst.nl 
'p.bosch@belastingdienst.nl      

'*rationale
'*De verwerking van een bronmodel naar een doelmodel is een tijdrovend en minutieus proces.
'*Daarbij mogen geen fouten worden gemaakt. Een oude wijsheid uit de systeemontwikkeling zegt dat fouten 
'*in het voortraject achteraf veel meer tijd kosten om te coorigeren dan fouten achter in het traject.
'*De ontwerpers van het EDW beschikken hiermee over een wizard die ondersteunt in hun werkzaamheden:
'*repetitief werk wordt verricht door tooling, terwijl ontwerkeuzes via stuurparameters door de ontwerpers aan 
'*de tooling wordt meegegeven. Hiermee bestaat een goede balans tussen geautomatiseerde en handmatig ontwerp.


'********** instructies **********
'1. start powerdesigner met het model geopend.
'2. open met ctrl+shift+x het edit/run script window en laad daarin dit script
'3. run het script met f5

'********** parameters **********
par_alternating_colors            = true
par_insert_empty_lines            = false
par_sort_referencejoins_on_column = false
par_underline_primary_keys        = true
par_gen_hide_worksheet            = true      'tabbladen verbergen
   
'>>> excel definitions  excel rows and columns are numbered 1, 2, 3, ...
dim excel_obj
dim matrix()
Dim words
excel_colorindex_nocolor = xlColorIndexNone'http://dmcritchie.mvps.org/excel/colors.htm
 excel_colorindex_black      = 1
 excel_colorindex_white      = 2
 excel_colorindex_red        = 3
 excel_colorindex_blue       = 5
 excel_colorindex_yellow     = 6
 excel_colorindex_magenta    = 7
 excel_colorindex_cyan       = 8
 excel_colorindex_green      = 10
 excel_colorindex_gray       = 15    
 excel_colorindex_bluegray   = 47
 excel_colorindex_source =excel_colorindex_yellow 
 excel_colorindex_target =excel_colorindex_bluegray 
 excel_colorindex_font_source = excel_colorindex_black
 excel_colorindex_font_target = excel_colorindex_white
 excel_colorindex_header = excel_colorindex_blue 
 excel_colorindex_avro = excel_colorindex_blue'
 excel_colorindex_font_header = excel_colorindex_white 
 excel_colorindex_1= excel_colorindex_black 
 excel_colorindex_2= excel_colorindex_blue
 zebra = excel_colorindex_1
 
'------------------------------------------------------------------------------ 
vbstartdate = date  
vbstarttime = time  

'------------------------------------------------------------------------------     
'main function  The programms starts here actually                                  
'------------------------------------------------------------------------------     

main
 
Output
Output "hope the run was ok!"
Output
output "VB Script started on       " & vbstartdate & " at " & vbstarttime

'------------------------------------------------------------------------------     
'main function  The programms ends here actually                                  
'------------------------------------------------------------------------------     

'In de map van de .pdm ontstaat een .xlsx spreadsheet met de zelfde naam (platform onafhankelijk PIM).
'Met behulp van  xls2pdm.vbs  kan deze spreadsheet weer worden geimporteerd in powerdesigner (platform specifiek PSM) 
'Advies gebruik detecteer_wijzigingen.vbs om vast te stellen dat er geen logische veschillen zijn ontstaan. 


private sub main
 
 '>>> definitions
 '>>> get the pdm filename and version
 pdm_filename = activemodel.filename
 do while instr(pdm_filename, "\") > 0
   pdm_filename = right(pdm_filename, len(pdm_filename) - instr(pdm_filename, "\"))
 loop
 
 bronsysteem =right(pdm_filename, 7 )
 bronsysteem =replace (bronsysteem , ".pdm", "" )
 pdm_version = activemodel.version
 
 '>>> create a new excel file
  set dummy = excel_new_file(false, 18) '18 onzichtbare werkbladen

    Set progressBar = Progress("Progress", true)
    progressBar.Min = 0
    progressBar.Max = 600 'Select the size of the progressbar. Usually you want to monitor progress through a collection iteration. Get its size.
    progressBar.Position = 0 'The bar does not need to start at 0, but remember that it will not move once (loop.count * bar.step) + bar.position > bar.max
    progressBar.Stepping = 1 'Unless you're doing some particular cases, you do not need to change the step size
    progressBar.Text = "pdm2xls is  running... " 'A text which will be displayed in the left part of the progress bar
    progressBar.BarColor = RGB(12,12,12)
    progressBar.Start() 'Display the progress bar

'*******************************************************************
 '>>> sheet 1: titel en versibeheer
 set dummy = excel_create_Titel_en_Historie 
 
'*******************************************************************
 '>>> sheet 2: entities
     'initialize the matrix
      set dummy = excel_select_worksheet(2)'entities
      progressBar.Step() 'Make the bar advance with progressBar.Stepping value        
      progressBar.Text = "pdm2xls is running entities"
     
      'initialize the matrix
       nr_entities = 0 
       nr_matrix_rows = 0
   for each table in activemodel.tables
     if not table.isshortcut then
       nr_matrix_rows = nr_matrix_rows + 1
      end if
       next
       nr_matrix_columns = 12
       redim matrix(nr_matrix_rows, nr_matrix_columns + 3)
       
      'fill the matrix and sort it
       matrix_row_nr = 0
       for each table in activemodel.tables
         if not table.isshortcut then
       	  nr_entities = nr_entities + 1
           
           progressBar.Step() 'Make the bar advance with progressBar.Stepping value        
           progressBar.Text = "pdm2xls is running entity:"&table.code  
           
       	  
     	   if table.annotation = "" then table.annotation =  ":>>:>>" end if 
       	  ' table.annotation =   "10:>>11:>>12"   'NB nog niet robuust. 
         
        If InStr( table.annotation, ":>>" ) > 0 Then 'the result should be true or false
                Words = Split(table.annotation, ":>>")
         end if         
         'output  words(0) &" <-> "& words(1) &"<->" & words(2)  &"<->" &table.code

    If InStr( table.annotation, ":>>") >  0 Then
           matrix(matrix_row_nr, 0)  = words(0)                                  'I     options                     1 
           matrix(matrix_row_nr, 1)  = words(1)                                  'U                                 2 
           matrix(matrix_row_nr, 2)  = words(2)                                  'D                                 3 
    end if       
           matrix(matrix_row_nr, 3)  = table.beginscript                         'USE-CASE                          4 
           matrix(matrix_row_nr, 4)  = table.name                                'ENTITY USE-CASE                   5 
           matrix(matrix_row_nr, 5)  = table.CheckConstraintName                 'FILE-NAME ACRONIEM                6 
           matrix(matrix_row_nr, 6)  = replace(table.description,Chr(10), " ")   'CALCULATION / TRANSFORMATION      7 
           matrix(matrix_row_nr, 7)  = table.endscript                           'ENVIRONMENT                       8 
           matrix(matrix_row_nr, 8)  = table.code                                'ENTITY                            9 
           matrix(matrix_row_nr, 9)  = table.stereotype                          'ENTITY STEREOTYPE                 10
           matrix(matrix_row_nr, 10) = replace(table.type,Chr(10), " ")          'CALCULATION / TRANSFORMATION      11
           matrix(matrix_row_nr, 11) = replace(table.comment,Chr(10), " ")       'ENTITY REMARKS                    12
                          
           matrix(matrix_row_nr, nr_matrix_columns + 0) = matrix(matrix_row_nr, 4) '4 = order by table.name
           matrix(matrix_row_nr, nr_matrix_columns + 2) = false
           matrix_row_nr = matrix_row_nr + 1
         end if
        next
       'apply zebra matrix coloring
       set dummy = matrix_sort_ascending(matrix, nr_matrix_rows, nr_matrix_columns)
        for matrix_row_nr = 0 to nr_matrix_rows - 1
         if matrix_row_nr mod 2 = 0 then
           matrix(matrix_row_nr, nr_matrix_columns + 1) = excel_colorindex_1
         else
            matrix(matrix_row_nr, nr_matrix_columns + 1) = excel_colorindex_2
         end if
       next
     'write the entity matrix to excel 
     set dummy = matrix_write_to_excel(matrix, nr_matrix_rows, nr_matrix_columns, 3, 1)
  
        
  '*******************************************************************
  '>>> sheet 3: attributes
      'initialize the matrix
      set dummy = excel_select_worksheet(3)
      progressBar.Step() 'Make the bar advance with progressBar.Stepping value        
      progressBar.Text = "pdm2xls is running attributes"
        
       'initialize the matrix     
       nr_attributes = 0
       nr_matrix_rows = 0
       for each table in activemodel.tables
      if not table.isshortcut then
       for each column in table.columns
      nr_matrix_rows = nr_matrix_rows + 1
       next
      end if
       next
       nr_matrix_columns = 24
       redim matrix(nr_matrix_rows, nr_matrix_columns + 3)
       
       'fill the matrix and sort it
       matrix_row_nr = 0
       for each table in activemodel.tables
       
       progressBar.Step() 'Make the bar advance with progressBar.Stepping value        
       progressBar.Text = "pdm2xls is running attribute :"&table.code  
       
        if not table.isshortcut then
          for each column in table.columns
            nr_attributes =  nr_attributes + 1 
            
      if column.annotation = "" then column.annotation =  ":>>:>>" end if 
            ' column.annotation =   "10:>>11:>>12"   'NB nog niet robuust.
       
       If InStr( column.annotation, ":>>") > 0 Then     
            Words = Split(column.annotation, ":>>")
            'output  words(0) &" <-> "& words(1) &"<->" & words(2)  &"<->" &column.code
       end if    

       If InStr( column.annotation, ":>>" ) > 0 Then
            matrix(matrix_row_nr, 0) =  words(0)                                                            'I                              1 
            matrix(matrix_row_nr, 1) =  words(1)                                                            'U                              2 
            matrix(matrix_row_nr, 2) =  words(2)                                                            'D                              3 
       end if     
            matrix(matrix_row_nr, 3) =  table.name                                                          'FILE                           4 
            matrix(matrix_row_nr, 4) =  column.name                                                         'FIELD                          5 
            
            matrix(matrix_row_nr, 5) =  column.format                                                       'FIELD-DATATYPE                 6 
               
            'vertaling
            if column.format = "" then
               matrix(matrix_row_nr, 5) =  column.datatype                                                  'FIELD-DATATYPE                 6    
            end if              
            
            matrix(matrix_row_nr, 6) =  column.physicaloptions                                              'FIELD-SEQUENCE                 7 
            matrix(matrix_row_nr, 7) =  column.lowvalue                                                     'FIELD-STARTPOSITION            8 
            matrix(matrix_row_nr, 8) =  column.highvalue                                                    'FIELD-ENDPOSITION              9 
            
            matrix(matrix_row_nr, 9) =  column.unit                                                         'FIELD-LENGTH                  10
           
           'vertaling
            if column.unit = "" then
            	matrix(matrix_row_nr, 9) =  column.length                                                     'FIELD-LENGTH                  10
            	if column.datatype = "INTEGER" then matrix(matrix_row_nr, 9) = 10 end if
            	if column.datatype = "SMALLINT" then matrix(matrix_row_nr, 9) = 4 end if
            	if column.datatype = "DATE" then matrix(matrix_row_nr, 9) = 10 end if
            	if column.datatype = "TIMESTAMP" then matrix(matrix_row_nr, 9) = 26 end if
            end if	
            
           If left(column.KeyIndicator,3) = "<ak" then matrix(matrix_row_nr, 10) = "AK" end if              'AK                            11
           if column.primary then matrix(matrix_row_nr, 10) = "PK" end if                                   'PK                            11
           if column.mandatory then matrix(matrix_row_nr, 11) = "X" end if                                  'MANDATORY                     12  
           if column.nospace then matrix(matrix_row_nr, 12) = "X" end if                                    'NOSPACE                       13   
            matrix(matrix_row_nr, 13) =  column.description                                                 'CALCULATION / TRANSFORMATION  14   
            matrix(matrix_row_nr, 14) =  table.endscript                                                    'ENVIRONMENT                   15   
            matrix(matrix_row_nr, 15) =  table.code                                                         'ENTITY                        16   
            matrix(matrix_row_nr, 16) =  column.code                                                        'ATTRIBUTE                     17   
            matrix(matrix_row_nr, 17) =  replace(column.datatype, Chr(10), " ")                             'DATA TYPE                     18   
            matrix(matrix_row_nr, 18) =  column.stereotype                                                  'ATTRIBUTE STEREOTYPE          19         
            matrix(matrix_row_nr, 19) =  column.ComputedExpression                                          'CALCULATION / TRANSFORMATION  20   
           if column.primary then matrix(matrix_row_nr, 20) = "PK" end if                                   'PK                            21
           if column.mandatory then matrix(matrix_row_nr, 21) = "X" end if                                  'MANDATORY                     22  
           if column.nospace then matrix(matrix_row_nr, 22) = "X" end if                                    'NOSPACE                       23   
           matrix(matrix_row_nr, 23) =  replace(column.comment , Chr(10), " ")                              'ATTRIBUTE REMARKS             24  
             
                                                                                                                                             
           matrix(matrix_row_nr, nr_matrix_columns + 0) = matrix(matrix_row_nr, 3) &"_" & (1000000 + matrix_row_nr)  ' order by 3 tablename
           matrix(matrix_row_nr, nr_matrix_columns + 2) = column.primary
           matrix_row_nr = matrix_row_nr + 1
         next
       end if
       next
       
    'apply zebra matrix coloring
       set dummy = matrix_sort_ascending(matrix, nr_matrix_rows, nr_matrix_columns)
        for matrix_row_nr = 0 to nr_matrix_rows - 1
          if matrix_row_nr = 0 then
             matrix(matrix_row_nr, nr_matrix_columns + 1) = excel_colorindex_1
         else
          if matrix(matrix_row_nr - 1, 3) <> matrix(matrix_row_nr, 3) then       '3 is sortcolumn 
            if matrix(matrix_row_nr - 1, nr_matrix_columns + 1) = excel_colorindex_2 then
            matrix(matrix_row_nr, nr_matrix_columns + 1) = excel_colorindex_1
           else
            matrix(matrix_row_nr, nr_matrix_columns + 1) = excel_colorindex_2
           end if
            else
           matrix(matrix_row_nr, nr_matrix_columns + 1) = matrix(matrix_row_nr - 1, nr_matrix_columns + 1)
            end if
           end if
       next
     'write the attribute matrix to excel   
      set dummy = matrix_write_to_excel(matrix, nr_matrix_rows, nr_matrix_columns, 3, 1)'  
       
 '*******************************************************************
 '>>> sheet 4: relations
     set dummy = excel_select_worksheet(4)
     progressBar.Step() 'Make the bar advance with progressBar.Stepping value        
      progressBar.Text = "pdm2xls is running relations"
    
     'initialize the matrix
     nr_relations = 0
     nr_matrix_rows = 0
     for each reference in activemodel.references
    if not reference.isshortcut then
    nr_matrix_rows = nr_matrix_rows + 1
    end if
     next
     nr_matrix_columns = 23
     redim matrix(nr_matrix_rows, nr_matrix_columns + 3)
   
     
     'fill the matrix and sort it
     matrix_row_nr = 0
     for each reference in activemodel.references
      if not reference.isshortcut then
      	 nr_relations = nr_relations + 1
          
         progressBar.Step() 'Make the bar advance with progressBar.Stepping value        
         progressBar.Text = "pdm2xls is running relation :"&reference.code
      	 
      	  if reference.annotation = "" then reference.annotation =  ":>>:>>" end if 
      	   'reference.annotation =   "10:>>11:>>12"   'NB nog niet robuust.
            
                    If InStr( ":>>", reference.annotation) > 0 Then     
              Words = Split(reference.annotation, ":>>")
               ' output  words(0) &" <-> "& words(1) &"<->" & words(2)  &"<->" &column.code
             end if 
                                    	 
         set parenttable = reference.parenttable
         set childtable = reference.childtable
               
       If InStr( ":>>", reference.annotation) > 0 Then 
         matrix(matrix_row_nr, 0  ) = words(0)                                                 '0   I                                   1                          
         matrix(matrix_row_nr, 1  ) = words(1)                                                 '1   U                                   2                          
         matrix(matrix_row_nr, 2  ) = words(2)                                                 '2   D                                   3                          
       end if 
         matrix(matrix_row_nr, 3  ) = reference.name                                           '3   name                                4  
         matrix(matrix_row_nr, 4  ) = reference.childtable.name                                '4   childtable.name                     5        
         matrix(matrix_row_nr, 5  ) = reference.parenttable.name                               '5   parenttable.name                    6        
         matrix(matrix_row_nr, 6  ) = reference.cardinality                                    '6   cardinality                         7      
        if reference.DeleteConstraint = 1  then matrix(matrix_row_nr, 7) = "X" end if          '7   Delete Constraint                   8    
        if reference.updateConstraint = 1  then matrix(matrix_row_nr, 8) = "X" end if          '8   Update Constraint                   9        
        if reference.ChangeParentAllowed then matrix(matrix_row_nr, 9)  = "X" end if           '9   ChangeParentAllowed                 10   
         matrix(matrix_row_nr, 10 ) = reference.childrole                                      '10  childrole                           11       
         matrix(matrix_row_nr, 11 ) = reference.parentrole                                     '11  parentrole                          12       
         matrix(matrix_row_nr, 12 ) = reference.comment                                        '12  reference.comment                   13       
         matrix(matrix_row_nr, 13 ) = reference.code                                           '13  reference.code                      14       
         matrix(matrix_row_nr, 14 ) = reference.childtable.code                                '14  reference.childtable.code           15       
         matrix(matrix_row_nr, 15 ) = reference.parenttable.code                               '15  reference.parenttable.code          16      
         matrix(matrix_row_nr, 16 ) = reference.stereotype                                     '16  reference.stereotype                17      
        if reference.ImplementationType  = "T"  then matrix(matrix_row_nr, 17) = "X" end if    '17  reference.ImplementationType        18  'maak een SAT onder de LNK      
         matrix(matrix_row_nr, 18 ) = reference.description                                    '18  reference.description               19      
         matrix(matrix_row_nr, 19 ) = replace(reference.JoinExpression, Chr(10), " " )         '19  reference.JoinExpression            20      
         matrix(matrix_row_nr, 20 ) = reference.ForeignKeyColumnList                           '20  reference.ForeignKeyColumnList      21      
         matrix(matrix_row_nr, 21 ) = reference.ParentKeyColumnList                            '21  reference.ParentKeyColumnList       22      
         matrix(matrix_row_nr, 22 ) = reference.foreignkeyconstraintname                       '22  reference.foreignkeyconstraintname  23      
                                                                                                         
         matrix(matrix_row_nr, nr_matrix_columns + 0) = matrix(matrix_row_nr, 4)  '4 = order by reference.name                                                                                        
         matrix(matrix_row_nr, nr_matrix_columns + 2) = false
         matrix_row_nr = matrix_row_nr + 1
       end if
     next
     
     'apply zebra matrix coloring
     set dummy = matrix_sort_ascending(matrix, nr_matrix_rows, nr_matrix_columns)
     for matrix_row_nr = 0 to nr_matrix_rows - 1
         if matrix_row_nr mod 2 = 0 then
          matrix(matrix_row_nr, nr_matrix_columns + 1) = excel_colorindex_1
         else
          matrix(matrix_row_nr, nr_matrix_columns + 1) = excel_colorindex_2
         end if
     next
     'write the relation matrix to excel
     set dummy = matrix_write_to_excel(matrix, nr_matrix_rows, nr_matrix_columns, 3, 1)
     
 '********************
 '>>> sheet 5: joins
     set dummy = excel_select_worksheet(5)
     progressBar.Step() 'Make the bar advance with progressBar.Stepping value        
      progressBar.Text = "pdm2xls is running joins"
    
    'initialize the matrix 
     nr_joins = 0
     nr_matrix_rows = 0
     for each reference in activemodel.references
    if not reference.isshortcut then
     for each referencejoin in reference.joins
    nr_matrix_rows = nr_matrix_rows + 1
     next
    end if
     next
     nr_matrix_columns = 10
     redim matrix(nr_matrix_rows, nr_matrix_columns + 3)
     
     'fill the matrix and sort it
     matrix_row_nr = 0
     for each reference in activemodel.references
      if not reference.isshortcut then
      
         progressBar.Step() 'Make the bar advance with progressBar.Stepping value        
         progressBar.Text = "pdm2xls is running relation :"&reference.code
      
      
        for each referencejoin in reference.joins
         
           nr_joins = nr_joins + 1
           matrix(matrix_row_nr, 0) = reference.name                                  ' reference.name                1 
           matrix(matrix_row_nr, 1) = reference.childtable.name                       ' reference.childtable.name     2 
           matrix(matrix_row_nr, 2) = referencejoin.childtablecolumn.name             ' join.childtablecolumn.name    3 
           matrix(matrix_row_nr, 3) = reference.parenttable.name                      ' reference.parenttable.name    4 
           matrix(matrix_row_nr, 4) = referencejoin.parenttablecolumn.name            ' join.parenttablecolumn.name   5 
           matrix(matrix_row_nr, 5) = reference.code                                  ' reference.code                6 
           matrix(matrix_row_nr, 6) = reference.childtable.code                       ' reference.childtable.code     7 
           matrix(matrix_row_nr, 7) = referencejoin.childtablecolumn.code             ' join.childtablecolumn.code    8 
           matrix(matrix_row_nr, 8) = reference.parenttable.code                      ' reference.parenttable.code    9 
           matrix(matrix_row_nr, 9) = referencejoin.parenttablecolumn.code            ' join.parenttablecolumn.code   10
           
           if par_sort_referencejoins_on_column then
             matrix(matrix_row_nr, nr_matrix_columns + 0) = matrix(matrix_row_nr, 1) & "_" & matrix(matrix_row_nr, 3) & "_" & matrix(matrix_row_nr, 2) & "_" & matrix(matrix_row_nr, 4)
           else
             matrix(matrix_row_nr, nr_matrix_columns + 0) = matrix(matrix_row_nr, 1) & "_" & matrix(matrix_row_nr, 3)
           end if
            matrix(matrix_row_nr, nr_matrix_columns + 2) = false
            matrix_row_nr = matrix_row_nr + 1
         next
        end if
     next
     
     'apply zebra matrix coloring
     set dummy = matrix_sort_ascending(matrix, nr_matrix_rows, nr_matrix_columns)
     for matrix_row_nr = 0 to nr_matrix_rows - 1
      if matrix_row_nr = 0 then
         matrix(matrix_row_nr, nr_matrix_columns + 1) = excel_colorindex_1
        else
         if (matrix(matrix_row_nr - 1, 1) & "_" & matrix(matrix_row_nr - 1, 3)) <> (matrix(matrix_row_nr, 1) & "_" & matrix(matrix_row_nr , 3)) then
        if matrix(matrix_row_nr - 1, nr_matrix_columns + 1) = excel_colorindex_2 then
         matrix(matrix_row_nr, nr_matrix_columns + 1) = excel_colorindex_1
        else
         matrix(matrix_row_nr, nr_matrix_columns + 1) = excel_colorindex_2
        end if
         else
        matrix(matrix_row_nr, nr_matrix_columns + 1) = matrix(matrix_row_nr - 1, nr_matrix_columns + 1)
         end if
      end if
     next
     'write the join matrix to excel
     set dummy = matrix_write_to_excel(matrix, nr_matrix_rows, nr_matrix_columns, 3, 1)

 '********************
 '>>> sheet 6: views
     set dummy = excel_select_worksheet(6)
     progressBar.Step() 'Make the bar advance with progressBar.Stepping value        
      progressBar.Text = "pdm2xls is running views"

     'initialize the matrix
     nr_matrix_rows = 0
     for each view in activemodel.views
        if not view.isshortcut then
         nr_matrix_rows = nr_matrix_rows + 1
        end if
     next
     nr_matrix_columns = 10
     redim matrix(nr_matrix_rows, nr_matrix_columns + 3)
     
    'fill the matrix and sort it
     matrix_row_nr = 0
     
   for each view in activemodel.views
    if not view.isshortcut then
    
         progressBar.Step() 'Make the bar advance with progressBar.Stepping value        
         progressBar.Text = "pdm2xls is running view :"&view.code 
    	
   	  if view.annotation = "" then view.annotation =  ":>>:>>" end if 
    	      view.annotation =   "10:>>11:>>12"   'NB nog niet robuust.
            Words = Split(view.annotation, ":>>")
            ' output  words(0) &" <-> "& words(1) &"<->" & words(2)  &"<->" &view.code
            matrix(matrix_row_nr, 0) = words(0)                                        ' I                         1
            matrix(matrix_row_nr, 1) = words(1)                                        ' U                         2
            matrix(matrix_row_nr, 2) = words(2)                                        ' D                         3
            matrix(matrix_row_nr, 3) = view.code                                       ' view.code                 4
            matrix(matrix_row_nr, 4) = view.name                                       ' view.name                 5
            matrix(matrix_row_nr, 5) = view.stereotype                                 ' view.stereotype           6
            matrix(matrix_row_nr, 6) = replace(view.comment, Chr(10), " ")             ' view.comment              7
            matrix(matrix_row_nr, 7) = view.type                                       ' view.type                 8
            matrix(matrix_row_nr, 8) = view.description                                ' view.description          9
            matrix(matrix_row_nr, 9) = replace(view.sqlquery,Chr(10), ":>>")           ' view.sqlquery             10
            matrix(matrix_row_nr, nr_matrix_columns + 0) = matrix(matrix_row_nr, 3)    '3 = order by view.code                                                                                        
            matrix(matrix_row_nr, nr_matrix_columns + 2) = false
            matrix_row_nr = matrix_row_nr + 1
     end if
    next

        set dummy = matrix_sort_ascending(matrix, nr_matrix_rows, nr_matrix_columns)
        for matrix_row_nr = 0 to nr_matrix_rows - 1
       if matrix_row_nr mod 2 = 0 then
          matrix(matrix_row_nr, nr_matrix_columns + 1) = excel_colorindex_1
       else
        matrix(matrix_row_nr, nr_matrix_columns + 1) = excel_colorindex_2
       end if
     next
     
     'write the view matrix to excel
       set dummy = matrix_write_to_excel(matrix, nr_matrix_rows, nr_matrix_columns, 3, 1)
       
 '******************** 
 '>>> sheet 7: viewcolumns
     set dummy = excel_select_worksheet(7)
     progressBar.Step() 'Make the bar advance with progressBar.Stepping value        
      progressBar.Text = "pdm2xls is running viewcolumns"
    
     'initialize the matrix
     nr_matrix_rows = 0
      for each view In ActiveModel.views
        if not view.isshortcut then
         for each column in view.columns
          nr_matrix_rows = nr_matrix_rows + 1
         next
       end if
     next
     
     nr_matrix_columns = 14
     redim matrix(nr_matrix_rows, nr_matrix_columns + 3)
     
     ' Dim words
     ' Words = Split(column.expression, ".")
     'output ">> :"& words(0) &" <-> "& words(1) &"->" & view.code & "_" &column.code & column.comment
      
      'fill the matrix and sort it
      matrix_row_nr = 0
      for each view in activemodel.views
        if not view.isshortcut then
        
         progressBar.Step() 'Make the bar advance with progressBar.Stepping value        
         progressBar.Text = "pdm2xls is running view :"&view.code 
        
          for each column in view.columns    
         

          if column.annotation = "" then column.annotation =  ":>>:>>" end if 
          'column.annotation =   "10:>>11:>>12"   'NB nog niet robuust.                  
          Words = Split(column.annotation, ":>>")                                       
          ' output  words(0) &" <-> "& words(1) &"<->" & words(2)  &"<->" &view.code   matrix(matrix_row_nr, 0) = column.annotation                              ' column.annotation             1 
           
           matrix(matrix_row_nr, 0) = words(0)                                       ' I                      2 
           matrix(matrix_row_nr, 1) = words(1)                                       ' U                      3 
           matrix(matrix_row_nr, 2) = words(2)                                       ' D                      4 
           matrix(matrix_row_nr, 3) = view.code                                      ' view.code              5            
           matrix(matrix_row_nr, 4) = column.code                                    ' column.code            5 
           matrix(matrix_row_nr, 5) = column.name                                    ' column.name            6 
           matrix(matrix_row_nr, 6) = column.format                                  ' column.format          7 
           matrix(matrix_row_nr, 7) = column.unit                                    ' len.                   8 
           matrix(matrix_row_nr, 8) = column.lowvalue                                ' pos_van.               9 
           matrix(matrix_row_nr, 9) = column.highvalue                               ' pos_tot                10
           matrix(matrix_row_nr, 10) = column.datatype                               ' column.mandatory       11
           matrix(matrix_row_nr, 11) = column.stereotype                             ' column.stereotype      12
           matrix(matrix_row_nr, 12) = replace(column.comment, Chr(10), " ")         ' column.comment         13
           matrix(matrix_row_nr, 13) = replace(column.description, Chr(10), " ")     ' column.description     14
                                                                                                                   '
           matrix(matrix_row_nr, nr_matrix_columns + 0) = matrix(matrix_row_nr, 3) &"_" & (1000000 + matrix_row_nr) '3 is sort column
           ' matrix(matrix_row_nr, nr_matrix_columns + 2) = column.primary
           matrix(matrix_row_nr, nr_matrix_columns + 2) = false
           matrix_row_nr = matrix_row_nr + 1
        next
       end if
     next
 'apply zebra matrix coloring
       set dummy = matrix_sort_ascending(matrix, nr_matrix_rows, nr_matrix_columns)
        for matrix_row_nr = 0 to nr_matrix_rows - 1
          if matrix_row_nr = 0 then
             matrix(matrix_row_nr, nr_matrix_columns + 1) = excel_colorindex_1
         else
          if matrix(matrix_row_nr - 1, 3) <> matrix(matrix_row_nr, 3) then       '3 is sortcolumn 
            if matrix(matrix_row_nr - 1, nr_matrix_columns + 1) = excel_colorindex_2 then
            matrix(matrix_row_nr, nr_matrix_columns + 1) = excel_colorindex_1
           else
            matrix(matrix_row_nr, nr_matrix_columns + 1) = excel_colorindex_2
           end if
            else
           matrix(matrix_row_nr, nr_matrix_columns + 1) = matrix(matrix_row_nr - 1, nr_matrix_columns + 1)
            end if
           end if
       next
      'write the viewcolumn matrix to excel
      set dummy = matrix_write_to_excel(matrix, nr_matrix_rows, nr_matrix_columns, 3, 1)
     
 '********************
 '>>> sheet 8: diagrams
     set dummy = excel_select_worksheet(8)
     progressBar.Step() 'Make the bar advance with progressBar.Stepping value        
      progressBar.Text = "pdm2xls is running diagram"
     
     'initialize the matrix
     nr_diagrams = 0
     nr_matrix_rows = 0
     for each diagram in activemodel.physicaldiagrams
       nr_matrix_rows = nr_matrix_rows + 1
     next
 
     nr_matrix_columns = 4
     redim matrix(nr_matrix_rows, nr_matrix_columns + 3)
     
     'fill the matrix and sort it
     matrix_row_nr = 0
     for each diagram in activemodel.physicaldiagrams
        nr_diagrams = nr_diagrams + 1
        
         progressBar.Step() 'Make the bar advance with progressBar.Stepping value        
         progressBar.Text = "pdm2xls is running diagram :"&diagram.code
        
        matrix(matrix_row_nr, 0) =   diagram.code
        matrix(matrix_row_nr, 1) =   diagram.name
        matrix(matrix_row_nr, 2) =   2 'diagram.pageformat      '2=landscape 1=portrait 
        matrix(matrix_row_nr, 3) =   132 'diagram.pageorientation '132 = A0
        
        matrix(matrix_row_nr, nr_matrix_columns + 0) = matrix(matrix_row_nr, 0)     '0 = order by diagram.code.
        matrix(matrix_row_nr, nr_matrix_columns + 2) = false
        matrix_row_nr = matrix_row_nr + 1
     next
     
     'apply zebra matrix coloring
     set dummy = matrix_sort_ascending(matrix, nr_matrix_rows, nr_matrix_columns)
     for matrix_row_nr = 0 to nr_matrix_rows - 1
         if matrix_row_nr mod 2 = 0 then
           matrix(matrix_row_nr, nr_matrix_columns + 1) = excel_colorindex_1
         else
          matrix(matrix_row_nr, nr_matrix_columns + 1) = excel_colorindex_2
       end if
     next
     'write the diagram matrix to excel
     set dummy = matrix_write_to_excel(matrix, nr_matrix_rows, nr_matrix_columns, 3, 1)
     
 '********************
 '>>> sheet 9: symbols
     'initialize the matrix
     set dummy = excel_select_worksheet(9)
     progressBar.Step() 'Make the bar advance with progressBar.Stepping value        
      progressBar.Text = "pdm2xls is running symbols"
     
     'initialize the matrix
     nr_symbols = 0
     nr_matrix_rows = 0
    for each diagram in activemodel.physicaldiagrams
     for each symbol in diagram.symbols
        if ((symbol.objecttype = "TableSymbol") or (symbol.objecttype = "ReferenceSymbol") or (symbol.objecttype = "ViewSymbol") ) then
       nr_matrix_rows = nr_matrix_rows + 1
        end if
     next
   next
     nr_matrix_columns = 16
     redim matrix(nr_matrix_rows, nr_matrix_columns + 3)
     
     'fill the matrix and sort it
     matrix_row_nr = 0
     for each diagram in activemodel.physicaldiagrams
     
     progressBar.Step() 'Make the bar advance with progressBar.Stepping value        
     progressBar.Text = "pdm2xls is running diagram :"&diagram.code
     
    for each symbol in diagram.symbols
    nr_symbols = nr_symbols + 1
    
    if ((symbol.objecttype = "TableSymbol") or (symbol.objecttype = "ReferenceSymbol") or (symbol.objecttype = "ViewSymbol")  ) then
          matrix(matrix_row_nr, 0) = diagram.code                                 ' diagram.code	1             
          matrix(matrix_row_nr, 1) = symbol.objecttype                            ' symbol.objecttype	2       
          matrix(matrix_row_nr, 2) = symbol.code                                  ' symbol.code	3             
          matrix(matrix_row_nr, 3) = symbol.name                                  ' symbol.name	4             
          matrix(matrix_row_nr, 4) = symbol.rect.top                              ' symbol.rect.top	5         
          matrix(matrix_row_nr, 5) = symbol.rect.bottom                           ' symbol.rect.bottom	6       
          matrix(matrix_row_nr, 6) = symbol.rect.left                             ' symbol.rect.left	7         
          matrix(matrix_row_nr, 7) = symbol.rect.right                            ' symbol.rect.right	8       
          matrix(matrix_row_nr, 8) = symbol.linecolor                             ' symbol.linecolor	9    
               
          if  (( symbol.objecttype = "TableSymbol") or (symbol.objecttype = "ViewSymbol") )  then                               ' 
              matrix(matrix_row_nr, 9)  = symbol.fillcolor                        ' symbol.fillcolor	10         
              matrix(matrix_row_nr, 10) = symbol.gradientfillmode                 ' symbol.gradientfillmode	11 
              matrix(matrix_row_nr, 11) = symbol.AutoAdjustToText                 ' symbol.AutoAdjustToText	12 
              matrix(matrix_row_nr, 12) = symbol.KeepAspect                       ' symbol.KeepAspect	13       
              matrix(matrix_row_nr, 13) = symbol.KeepCenter                       ' symbol.KeepCenter	14       
              matrix(matrix_row_nr, 14) = symbol.KeepSize                         ' symbol.KeepSize	15         
          end if
          
          if symbol.objecttype = "TableSymbol"       then           matrix(matrix_row_nr, 15) =  1  end if
          if symbol.objecttype = "ReferenceSymbol"   then           matrix(matrix_row_nr, 15) =  2  end if
          if symbol.objecttype = "ViewSymbol"        then           matrix(matrix_row_nr, 15) =  3  end if

           matrix(matrix_row_nr, nr_matrix_columns + 0) = matrix(matrix_row_nr, 15)    '15 = order by symbol.objecttype  views komen onderaan.
           matrix(matrix_row_nr, nr_matrix_columns + 2) = false                                                        
           matrix_row_nr = matrix_row_nr + 1                                                                           

     end if
     next                                               
     
    next
    'apply zebra matrix coloring
    
     set dummy = matrix_sort_ascending(matrix, nr_matrix_rows, nr_matrix_columns)
     for matrix_row_nr = 0 to nr_matrix_rows - 1
       if matrix_row_nr = 0 then
         matrix(matrix_row_nr, nr_matrix_columns + 1) = excel_colorindex_1
        else
         if matrix(matrix_row_nr, 0) <> matrix(matrix_row_nr - 1, 15) then  '1,15=sortkolumn
        if matrix(matrix_row_nr - 1, nr_matrix_columns + 1) = excel_colorindex_2 then
         matrix(matrix_row_nr, nr_matrix_columns + 1) = excel_colorindex_1
        else
         matrix(matrix_row_nr, nr_matrix_columns + 1) = excel_colorindex_2
        end if
         else
        matrix(matrix_row_nr, nr_matrix_columns + 1) = matrix(matrix_row_nr - 1, nr_matrix_columns + 1)
        end if
      end if
     next
     'write the symbol matrix to excel
      set dummy = matrix_write_to_excel(matrix, nr_matrix_rows, nr_matrix_columns, 3, 1)
     
'set headers 

         progressBar.Step() 'Make the bar advance with progressBar.Stepping value        
         progressBar.Text = "pdm2xls is running headers"

set dummy = excel_create_headers4LGM

'*******************************************************************
 '>>> sheet 10: 
  set dummy = excel_select_worksheet(10)
  set dummy = excel_set_worksheet_name(10, "sheet10" , true )
  set dummy = excel_create_sheet10 
'*******************************************************************
 '>>> sheet 11: 
 'set dummy = excel_create_sheet10 
  set dummy = excel_select_worksheet(11)
  set dummy = excel_set_worksheet_name(11, "sheet11" , true )
  set dummy = excel_create_sheet11 
'*******************************************************************

 '>>> sheet 12: 
 'set dummy = excel_create_sheet12 
  set dummy = excel_select_worksheet(12)
  set dummy = excel_set_worksheet_name(12, "sheet12" , true )
  set dummy = excel_create_sheet12

'*******************************************************************
 '>>> sheet 13: 
 'set dummy = excel_create_sheet13 
  set dummy = excel_select_worksheet(13)
  set dummy = excel_set_worksheet_name(13, "sheet13" , true )
  set dummy = excel_create_sheet13
'*******************************************************************

 '>>> sheet 14: 
 'set dummy = excel_create_sheet14 
  set dummy = excel_select_worksheet(14)
  set dummy = excel_set_worksheet_name(14, "sheet14" , true )
  set dummy = excel_create_sheet14
'*******************************************************************
 '>>> sheet 15: 
 'set dummy = excel_create_sheet15 
 
  set dummy = excel_select_worksheet(15)
  set dummy = excel_set_worksheet_name(15, "sheet15" , true )
  set dummy = excel_create_sheet15

'*******************************************************************
 '>>> sheet 16: 
 'set dummy = excel_create_sheet16 
  set dummy = excel_select_worksheet(16)
  set dummy = excel_set_worksheet_name(16, "sheet16" , true )
  set dummy = excel_create_sheet16
'*******************************************************************

'*******************************************************************
 '>>> sheet 17: view references
   ' for each vwref in activemodel.viewreferences
'    if not vwref.isshortcut then
'     output vwref.code
'    end if
'next    

     set dummy = excel_select_worksheet(17)
     progressBar.Step() 'Make the bar advance with progressBar.Stepping value        
      progressBar.Text = "pdm2xls is running view viewreferences"
    
     'initialize the matrix
     nr_vwref = 0
     nr_matrix_rows = 0
     for each vwref in activemodel.viewreferences
    if not vwref.isshortcut then
    nr_matrix_rows = nr_matrix_rows + 1
    end if
     next
     nr_matrix_columns = 23
     redim matrix(nr_matrix_rows, nr_matrix_columns + 3)
     
     'fill the matrix and sort it
     matrix_row_nr = 0
     for each vwref in activemodel.viewreferences
      if not vwref.isshortcut then
      	 nr_vwref = nr_vwref + 1
          
         progressBar.Step() 'Make the bar advance with progressBar.Stepping value        
         progressBar.Text = "pdm2xls is running relation :"&vwref.code
               
         matrix(matrix_row_nr, 0 ) = "tt"                                       ' I                                          1                          
         matrix(matrix_row_nr, 1 ) = "tt"                                       ' U                                          2                          
         matrix(matrix_row_nr, 2 ) = "tt"                                       ' D                                          3                          
         matrix(matrix_row_nr, 3 ) =  vwref.name                                '3   name                                    4  
         matrix(matrix_row_nr, 4 ) = "tt"                                       '4   childtable.name                         5        
         matrix(matrix_row_nr, 5 ) = "tt"                                       '5   parenttable.name                        6        
         matrix(matrix_row_nr, 6 ) = "tt"                                       '6   cardinality                             7      
         matrix(matrix_row_nr, 7 ) = "tt"                                       '7   Delete Constraint                       8    
         matrix(matrix_row_nr, 8 ) = "tt"                                       '8   Update Constraint                       9        
         matrix(matrix_row_nr, 9 ) = "tt"                                       '9  ChangeParentAllowed                    10   
         matrix(matrix_row_nr, 10) = "tt"                                       '10  childrole                               11       
         matrix(matrix_row_nr, 11) = "tt"                                       '11  parentrole                              12       
         matrix(matrix_row_nr, 12) = "tt"                                       '12  viewreference.comment                   13       
         matrix(matrix_row_nr, 13) =  vwref.code                                '13  viewreference.code                      14       
         matrix(matrix_row_nr, 14) =  vwref.tableview2.code                     '14  viewreference.childtable.code           15       
         matrix(matrix_row_nr, 15) =  vwref.tableview1.code                     '15  viewreference.parenttable.code          16      
         matrix(matrix_row_nr, 16) = "tt"                                       '16  viewreference.stereotype                17      
         matrix(matrix_row_nr, 17) = "tt"                                       '17  viewreference.ImplementationType        18 
         matrix(matrix_row_nr, 18) = "tt"                                       '18  viewreference.description               19      
         matrix(matrix_row_nr, 19) = "tt"                                       '19  viewreference.JoinExpression            20      
         matrix(matrix_row_nr, 20) = "tt"                                       '20  viewreference.ForeignKeyColumnList      21      
         matrix(matrix_row_nr, 21) = "tt"                                       '21  viewreference.ParentKeyColumnList       22      
         matrix(matrix_row_nr, 22) = "tt"                                       '22  viewreference.foreignkeyconstraintname  23      
                                                                                                         
         matrix(matrix_row_nr, nr_matrix_columns + 0) = matrix(matrix_row_nr, 4)  '4 = order by viewreference.name                                                                                        
         matrix(matrix_row_nr, nr_matrix_columns + 2) = false
         matrix_row_nr = matrix_row_nr + 1
       end if
     next
     
     'apply zebra matrix coloring
     set dummy = matrix_sort_ascending(matrix, nr_matrix_rows, nr_matrix_columns)
     for matrix_row_nr = 0 to nr_matrix_rows - 1
         if matrix_row_nr mod 2 = 0 then
          matrix(matrix_row_nr, nr_matrix_columns + 1) = excel_colorindex_1
         else
          matrix(matrix_row_nr, nr_matrix_columns + 1) = excel_colorindex_2
         end if
     next
     'write the relation matrix to excel
     set dummy = matrix_write_to_excel(matrix, nr_matrix_rows, nr_matrix_columns, 3, 1)
     
 '********************
 '>>> sheet 18: viewreference joins
     set dummy = excel_select_worksheet(18)
     progressBar.Step() 'Make the bar advance with progressBar.Stepping value        
      progressBar.Text = "pdm2xls is running viewreference joins"
    
    'initialize the matrix 
     nr_joins = 0
     nr_matrix_rows = 0
     for each reference in activemodel.references
    if not reference.isshortcut then
     for each referencejoin in reference.joins
    nr_matrix_rows = nr_matrix_rows + 1
     next
    end if
     next
     nr_matrix_columns = 10
     redim matrix(nr_matrix_rows, nr_matrix_columns + 3)
     
     'fill the matrix and sort it
     matrix_row_nr = 0
     for each reference in activemodel.references
      if not reference.isshortcut then
      
         progressBar.Step() 'Make the bar advance with progressBar.Stepping value        
         progressBar.Text = "pdm2xls is running viewreference join :"&reference.code
      
        for each referencejoin in reference.joins
         
           nr_joins = nr_joins + 1
           matrix(matrix_row_nr, 0) = reference.name                                  ' reference.name                1 
           matrix(matrix_row_nr, 1) = reference.childtable.name                       ' reference.childtable.name     2 
           matrix(matrix_row_nr, 2) = referencejoin.childtablecolumn.name             ' join.childtablecolumn.name    3 
           matrix(matrix_row_nr, 3) = reference.parenttable.name                      ' reference.parenttable.name    4 
           matrix(matrix_row_nr, 4) = referencejoin.parenttablecolumn.name            ' join.parenttablecolumn.name   5 
           matrix(matrix_row_nr, 5) = reference.code                                  ' reference.code                6 
           matrix(matrix_row_nr, 6) = reference.childtable.code                       ' reference.childtable.code     7 
           matrix(matrix_row_nr, 7) = referencejoin.childtablecolumn.code             ' join.childtablecolumn.code    8 
           matrix(matrix_row_nr, 8) = reference.parenttable.code                      ' reference.parenttable.code    9 
           matrix(matrix_row_nr, 9) = referencejoin.parenttablecolumn.code            ' join.parenttablecolumn.code   10
           
           if par_sort_referencejoins_on_column then
             matrix(matrix_row_nr, nr_matrix_columns + 0) = matrix(matrix_row_nr, 1) & "_" & matrix(matrix_row_nr, 3) & "_" & matrix(matrix_row_nr, 2) & "_" & matrix(matrix_row_nr, 4)
           else
             matrix(matrix_row_nr, nr_matrix_columns + 0) = matrix(matrix_row_nr, 1) & "_" & matrix(matrix_row_nr, 3)
           end if
            matrix(matrix_row_nr, nr_matrix_columns + 2) = false
            matrix_row_nr = matrix_row_nr + 1
         next
        end if
     next
     
     'apply zebra matrix coloring
     set dummy = matrix_sort_ascending(matrix, nr_matrix_rows, nr_matrix_columns)
     for matrix_row_nr = 0 to nr_matrix_rows - 1
      if matrix_row_nr = 0 then
         matrix(matrix_row_nr, nr_matrix_columns + 1) = excel_colorindex_1
        else
         if (matrix(matrix_row_nr - 1, 1) & "_" & matrix(matrix_row_nr - 1, 3)) <> (matrix(matrix_row_nr, 1) & "_" & matrix(matrix_row_nr , 3)) then
        if matrix(matrix_row_nr - 1, nr_matrix_columns + 1) = excel_colorindex_2 then
         matrix(matrix_row_nr, nr_matrix_columns + 1) = excel_colorindex_1
        else
         matrix(matrix_row_nr, nr_matrix_columns + 1) = excel_colorindex_2
        end if
         else
        matrix(matrix_row_nr, nr_matrix_columns + 1) = matrix(matrix_row_nr - 1, nr_matrix_columns + 1)
         end if
      end if
     next
     'write the join matrix to excel
     set dummy = matrix_write_to_excel(matrix, nr_matrix_rows, nr_matrix_columns, 3, 1)

  '>>> hiding 
    
    if par_gen_hide_worksheet = true then 
  	 set dummy =excel_hide_columns
  	 set dummy = excel_hide_worksheets
  end if '	 
  

'klaar met verwerking         
'>>> close the excel file
        
      
set dummy = excel_select_worksheet(1) 
 
' file_name = Replace(left(activemodel.filename, len(activemodel.filename) - 4) & ".xls", ":\", ":\interfacebeschrijving_")
'output file_name 
'set dummy = excel_save_as_file(file_name)

nu= year(date) & month (date) & day(date) & "_"  & hour(time)& "u" & minute(time) & "_"  & second(time)   
file_name = left(activemodel.filename, len(activemodel.filename) - 4)&"_"&pdm_version      
file_name =replace (file_name , ".", "_" ) 'even filename optimaliseren
file_name =replace (file_name , " ", "_")
file_name =replace (file_name , "&", "_")

set dummy = excel_save_as_file(file_name&".xlsx")
set dummy = excel_close_file()
 
 output  "'************   "& file_name
 output  "entiteiten     = "&nr_entities
 output  "attributen     = "&nr_attributes
 output  "relaties       = "&nr_relations
 output  "joins          = "&nr_joins
 output  "diagrammen     = "&nr_diagrams
 output  "symbolen       = "&nr_symbols
 output  "'************ " & file_name
 
 
 progressBar.Text = "hope the run was okay!."
'progressBar.Stop() 'Hide the progress bar

 
  nu= year(date) & month (date) & day(date) & "_"  & hour(time)& "u" & minute(time) & "_"  & second(time)                    
  'par_outfile = "c:\temp\08205_LGM_AGT_STI_ICT_v1_"&nu&".xlsx"     'Setter PIM Platform Independant Model      
  output  "" 
  'RQ = MsgBox ("Is Excel Installed on your machine ?", vbYesNo + vbInformation,"Confirmation")
  output "output is weggeschreven naar = "& file_name & ".xlsx (use xls2pdm  for  import PSM  platform specic model in powerdesigner) "
  
  
  
  
  
end sub


'end of processing 

'------------------------------------------------------------------------------
'#Powerdesigner is te beschouwen als een [ObjectOriented XML database] 
'#De OO getters setters staan hieronder
'------------------------------------------------

'>>> functions that are called by other functions
'# Structured programming:
'# Entire program logic modularized in functions.
' ------------------------------
 
'********** excel functions **********
' many thanks to cc.wust@belastingdienst.nl
'>>> functions that are called by other functions
'excel rows and columns are numbered 1, 2, 3, ...
'>>> definitions
 
function excel_get_row(excel_row_nr)
  excel_get_row = cstr(excel_row_nr)
end function
 
function excel_get_rows(excel_row_nr1, excel_row_nr2)
 excel_get_rows = excel_get_row(excel_row_nr1) & ":" & excel_get_row(excel_row_nr2)
end function
 
function excel_get_column(excel_column_nr)
    letter1 = (excel_column_nr - 1) \ 26
   letter2 = ((excel_column_nr - 1) mod 26) + 1
 if letter1 = 0 then
   excel_get_column = chr(64 + letter2)
 else
excel_get_column = chr(64 + letter1) & chr(64 + letter2)
 end if
end function
 
function excel_get_columns(excel_column_nr1, excel_column_nr2)
 excel_get_columns = excel_get_column(excel_column_nr1) & ":" & excel_get_column(excel_column_nr2)
end function
 
function excel_get_cell(excel_row_nr, excel_column_nr)
 excel_get_cell = excel_get_column(excel_column_nr) & excel_get_row(excel_row_nr)
end function
 
function excel_get_cells(excel_row_nr1, excel_column_nr1, excel_row_nr2, excel_column_nr2)
 excel_get_cells = excel_get_column(excel_column_nr1) & excel_get_row(excel_row_nr1) & ":" & excel_get_column(excel_column_nr2) & excel_get_row(excel_row_nr2)
end function
 
'>>> functions to set or read the value of a cell
 
function excel_set_cell_value(excel_row_nr, excel_column_nr, value)
 excel_obj.range(excel_get_cell(excel_row_nr, excel_column_nr)).value = value
 set excel_set_cell_value = nothing
end function
 
 
function excel_set_cell_comment(excel_row_nr, excel_column_nr, value)
 
set dummy = excel_set_cursor(excel_row_nr, excel_column_nr)
   With excel_obj.range(excel_get_cell(excel_row_nr, excel_column_nr))
    .AddComment  value
     
     With .Comment.Shape
        .Height = 200
        .Width = 100
    End With
'    .Comment.Visible = false
 
   End With 
   set excel_set_cell_comment = nothing
end function
 
function excel_get_cell_value(excel_row_nr, excel_column_nr)
 excel_get_cell_value = excel_obj.range(excel_get_cell(excel_row_nr, excel_column_nr)).value
end function
 
'>>> i/o functions
 
function excel_new_file(visible, nr_sheets)
 set excel_obj = createobject("excel.application")
 excel_obj.visible = visible
 excel_obj.sheetsinnewworkbook = nr_sheets
 excel_obj.workbooks.add
 set excel_new_file = nothing
end function
 
function excel_load_file(visible, filename)
 set excel_obj = createobject("excel.application")
 excel_obj.visible = visible
 excel_obj.workbooks.open(filename)
 excel_obj.worksheets(1).select
 set excel_load_file = nothing
end function
 
function excel_save_file()
 excel_obj.activeworkbook.save
 set excel_save_file = nothing
end function
 
function excel_save_as_file(filename)
 excel_obj.activeworkbook.saveas(filename)
 
 set excel_save_as_file = nothing
end function
 
function excel_close_file()
 excel_obj.activeworkbook.close
 set excel_close_file = nothing
end function
 
'>>> worksheet functions
 
function excel_select_worksheet(worksheet_nr)
 excel_obj.worksheets(worksheet_nr).select
 set excel_select_worksheet = nothing
end function
 
function excel_set_worksheet_name(worksheet_nr, worksheet_name, visible )
 excel_obj.worksheets(worksheet_nr).name = worksheet_name
 excel_obj.worksheets(worksheet_nr).visible = visible
 
set excel_set_worksheet_name = nothing
end function
 
function excel_get_worksheet_name(worksheet_nr)
 excel_get_worksheet_name = excel_obj.worksheets(worksheet_nr).name
end function

function excel_set_worksheet_tabcolor(worksheet_nr, tabkleur )
   excel_obj.worksheets(worksheet_nr).tab.colorindex  = tabkleur
   set excel_set_worksheet_tabcolor = nothing
end function

 
'>>> layout functions
 
 function  excel_set_autofilter(excel_row_nr1, excel_column_nr1, excel_row_nr2, excel_column_nr2)
   excel_obj.range(excel_get_cells(excel_row_nr1, excel_column_nr1, excel_row_nr2, excel_column_nr2)).AutoFilter
       set excel_set_autofilter = nothing
end function
 
function excel_set_cursor(excel_row_nr, excel_column_nr)
 excel_obj.range(excel_get_cell(excel_row_nr, excel_column_nr)).select
 set excel_set_cursor = nothing
end function
 
 
function excel_merge_cells(excel_row_nr1, excel_column_nr1, excel_row_nr2, excel_column_nr2)
 excel_obj.range(excel_get_cells(excel_row_nr1, excel_column_nr1, excel_row_nr2, excel_column_nr2)).select
 excel_obj.selection.merge
 set excel_merge_cells = nothing
end function
 
function excel_set_zoom(zoom_percentage)
 excel_obj.activewindow.zoom = zoom_percentage
 set excel_set_zoom = nothing
end function
 
function excel_autofit_columns(excel_column_nr1, excel_column_nr2)
 excel_obj.columns(excel_get_columns(excel_column_nr1, excel_column_nr2)).autofit
 set excel_autofit_columns = nothing
end function
 
function excel_freeze_row(excel_row_nr)
 excel_obj.rows(excel_get_rows(excel_row_nr, excel_row_nr)).select
 excel_obj.activewindow.freezepanes = true
 set excel_freeze_row = nothing
end function
 
function excel_freeze_column(excel_column_nr)
 excel_obj.columns(excel_get_columns(excel_column_nr, excel_column_nr)).select
 excel_obj.activewindow.freezepanes = true
 set excel_freeze_column = nothing
end function
 
function excel_freeze_cell(excel_row_nr, excel_column_nr)
 excel_obj.range(excel_get_cell(excel_row_nr, excel_column_nr)).select
 excel_obj.activewindow.freezepanes = true
 set excel_freeze_cell = nothing
end function
 
function excel_set_cells_background_color(excel_row_nr1, excel_column_nr1, excel_row_nr2, excel_column_nr2, excel_colorindex)
 excel_obj.range(excel_get_cells(excel_row_nr1, excel_column_nr1, excel_row_nr2, excel_column_nr2)).interior.colorindex = excel_colorindex
 set excel_set_cells_background_color = nothing
end function
 
function excel_set_cells_border_color(excel_row_nr1, excel_column_nr1, excel_row_nr2, excel_column_nr2, excel_colorindex)
 excel_obj.range(excel_get_cells(excel_row_nr1, excel_column_nr1, excel_row_nr2, excel_column_nr2)).select
 excel_obj.selection.borders.linestyle = xlContinuous
excel_obj.selection.borders.colorindex = excel_colorindex
 
 set excel_set_cells_border_color = nothing
end function

function excel_set_font_columns(excel_column_nr1, excel_column_nr2)
' excel_obj.columns(excel_get_columns(excel_column_nr1, excel_column_nr2)).font.Name = "Arial"
 excel_obj.columns(excel_get_columns(excel_column_nr1, excel_column_nr2)).font.Name = "Lucida Console"
 excel_obj.columns(excel_get_columns(excel_column_nr1, excel_column_nr2)).font.size = 10
 set excel_set_font_columns = nothing
end function
 
function excel_set_cells_font_color(excel_row_nr1, excel_column_nr1, excel_row_nr2, excel_column_nr2, excel_colorindex)
 excel_obj.range(excel_get_cells(excel_row_nr1, excel_column_nr1, excel_row_nr2, excel_column_nr2)).font.colorindex = excel_colorindex
 set excel_set_cells_font_color = nothing
end function
 
function excel_set_cells_bold(excel_row_nr1, excel_column_nr1, excel_row_nr2, excel_column_nr2, activate)
 excel_obj.range(excel_get_cells(excel_row_nr1, excel_column_nr1, excel_row_nr2, excel_column_nr2)).font.bold = activate
 set excel_set_cells_bold = nothing
end function
 
function excel_set_cells_italic(excel_row_nr1, excel_column_nr1, excel_row_nr2, excel_column_nr2, activate)
 excel_obj.range(excel_get_cells(excel_row_nr1, excel_column_nr1, excel_row_nr2, excel_column_nr2)).font.italic = activate
 set excel_set_cells_italic = nothing
end function
 
function excel_set_cells_underline(excel_row_nr1, excel_column_nr1, excel_row_nr2, excel_column_nr2, activate)
 excel_obj.range(excel_get_cells(excel_row_nr1, excel_column_nr1, excel_row_nr2, excel_column_nr2)).font.underline = activate
 set excel_set_cells_underline = nothing
end function                    

'=======================================================================================

function excel_create_Titel_en_Historie()  
 '>>> sheet 1: Titel en Historie   
 
    'Bepaal parameters van het model  
    'Algemene vulling excel
    
    set dummy = excel_select_worksheet(1)

    set dummy = excel_set_worksheet_name(1, "Titel + versiebeheer", true)
    set dummy = excel_set_cell_value(1, 1, pdm_version)
    set dummy = excel_set_cell_value(1, 2, pdm_filename)
    
    set dummy = excel_set_cell_value(2, 1, "Belastingdienst")
    set dummy = excel_set_cell_value(2, 4, " Status: "& pdm_version )
    set dummy = excel_set_cell_value(2, 5, date & " - " & time )
    
    set dummy = excel_set_cells_bold(1, 1, 1, 2, true)
    set dummy = excel_set_cells_background_color(1, 1, 1, 5, excel_colorindex_avro)
    set dummy = excel_set_cells_background_color(2, 1, 2, 5, excel_colorindex_header)
    set dummy = excel_set_cells_font_color(1, 1, 2, 6, excel_colorindex_font_header)
   
    'Wegschrijven annotation naar excel
    'Wegschrijven headers parameters
    set dummy = excel_set_cell_value(4, 1, "PARAMETERS")
    set dummy = excel_set_cell_value(5, 1, "(Structuur en nummering niet aanpassen !!)")
    
    'Wegschrijven detail parameters
    set dummy = excel_set_cell_value(7, 1, "1.EDW Domein		 ")
    set dummy = excel_set_cell_value(8, 1, "2.EDW Recordbronnaam	 ")
    set dummy = excel_set_cell_value(9, 1, "3.EDW Volgnummer	 ")
    set dummy = excel_set_cell_value(10, 1, "4.EDW Type ontsluiting	")
    set dummy = excel_set_cell_value(11, 1, "5.EDW Prefix		 ")
    
    'wegschrijven headers versiebeheer
    set dummy = excel_set_cell_value(17, 1, "VERSIEBEHEER" )
    set dummy = excel_set_cell_value(18, 1, "(Structuur niet aanpassen !!)")

    set dummy = excel_set_cells_background_color(19, 1, 19, 6, excel_colorindex_gray) 
    set dummy = excel_set_cells_border_color(19, 1, 19, 6, excel_colorindex_black) 

    
    'Indien er vulling is velden in excel toeveogen
    If activemodel.annotation <> "" and instr(activemodel.annotation,"1.EDW") > 1 then 
    
    MDL_annotation= MID(activemodel.annotation,instr(activemodel.annotation,"1.EDW"),500) 
    'output MDL_annotation
    
    v_EDW_tmp = MID(MDL_annotation,instr(MDL_annotation,"1.EDW"),500) 
    v_EDW_Domein = mid(v_EDW_tmp, instr(v_EDW_tmp,":")+2 ,instr(v_EDW_tmp,"\par")-instr(v_EDW_tmp,":")-2) 
    v_EDW_DOMEIN = Replace(v_EDW_Domein, Chr(10), " " ) 
    v_EDW_DOMEIN = Replace(v_EDW_Domein, Chr(13), " " ) 
    v_EDW_tmp = MID(MDL_annotation,instr(MDL_annotation,"2.EDW"),500) 
    v_EDW_Recordbron = mid(v_EDW_tmp, instr(v_EDW_tmp,":")+2 ,instr(v_EDW_tmp,"\par")-instr(v_EDW_tmp,":")-2) 
    v_EDW_Recordbron = Replace(v_EDW_Recordbron, Chr(10), " " ) 
    v_EDW_Recordbron = Replace(v_EDW_Recordbron, Chr(13), " " )
    v_EDW_Recordbron = trim(v_EDW_Recordbron) 
    v_EDW_tmp = MID(MDL_annotation,instr(MDL_annotation,"3.EDW"),500) 
    v_EDW_Volgnummer = mid(v_EDW_tmp, instr(v_EDW_tmp,":")+2 ,instr(v_EDW_tmp,"\par")-instr(v_EDW_tmp,":")-2) 
    v_EDW_Volgnummer = Replace(v_EDW_Volgnummer, Chr(10), " " ) 
    v_EDW_Volgnummer = "'" & Replace(v_EDW_Volgnummer, Chr(13), " " )
     
    v_EDW_tmp = MID(MDL_annotation,instr(MDL_annotation,"4.EDW"),500) 
    v_EDW_Type_ontsl = mid(v_EDW_tmp, instr(v_EDW_tmp,":")+2 ,instr(v_EDW_tmp,"\par")-instr(v_EDW_tmp,":")-2) 
    v_EDW_Type_ontsl = Replace(v_EDW_Type_ontsl, Chr(10), " " ) 
    v_EDW_Type_ontsl = Replace(v_EDW_Type_ontsl, Chr(13), " " )
    v_EDW_Type_ontsl = trim(v_EDW_Type_ontsl) 
    v_EDW_tmp = MID(MDL_annotation,instr(MDL_annotation,"5.EDW"),500) 
    v_EDW_Prefix = mid(v_EDW_tmp, instr(v_EDW_tmp,":")+2 ,2)
 
    'output "MDL_annotation = "& MDL_annotation  
    output "v_EDW_Domein = " & v_EDW_Domein
    output "v_EDW_Recordbron= " & v_EDW_Recordbron
    output "v_EDW_Volgnummer= "& v_EDW_Volgnummer
    output "v_EDW_Type_ontsl= " & v_EDW_Type_ontsl
    output "v_EDW_Prefix= " & v_EDW_Prefix
 
    set dummy = excel_set_cell_value(7, 2, v_EDW_Domein)
    set dummy = excel_set_cell_value(8, 2, v_EDW_Recordbron)
    set dummy = excel_set_cell_value(9, 2, v_EDW_Volgnummer)
    set dummy = excel_set_cell_value(10, 2, v_EDW_Type_ontsl)
    set dummy = excel_set_cell_value(11, 2, v_EDW_Prefix)
    end if
    
    
    If activemodel.description <> ""  and instr(activemodel.description,"Versie") > 1 then 
           'wegschrijven details versiebeheer
          MDL_desc = mid(activemodel.description,instr(activemodel.description,"Versie"),10000) 
          MDL_desc = "||" & MDL_desc & " ||"
          MDL_desc = replace(MDL_desc,chr(10),"||")
          MDL_desc = replace(MDL_desc,chr(9)," ")
          MDL_desc = replace(MDL_desc,chr(13)," ")
          MDL_desc = replace(MDL_desc,"\par"," ")
          MDL_desc = replace(MDL_desc,"}"," ")
          MDL_desc = replace(MDL_desc,"\tab"," ")
 
          p=1 
          Do While p < 50
             MDL_desc = MID(MDL_desc,instr(MDL_desc,"||")+2,10000)
            ' output "MDL_desc= " & MDL_desc 
             If replace(replace(MDL_desc,"|","")," ","") <> "" then 
                MDL_desc_versie = MID(MDL_desc,1,instr(MDL_desc,"|")-1)
                MDL_desc = MID(MDL_desc,instr(MDL_desc,"|")+1,10000)
                MDL_desc_basis = MID(MDL_desc,1,instr(MDL_desc,"|")-1)
                MDL_desc = MID(MDL_desc,instr(MDL_desc,"|")+1,10000)
                MDL_desc_datum = MID(MDL_desc,1,instr(MDL_desc,"|")-1)
                MDL_desc = MID(MDL_desc,instr(MDL_desc,"|")+1,10000)
                MDL_desc_auteur = MID(MDL_desc,1,instr(MDL_desc,"|")-1)
                MDL_desc = MID(MDL_desc,instr(MDL_desc,"|")+1,10000)
                MDL_desc_taak = MID(MDL_desc,1,instr(MDL_desc,"|")-1)
                MDL_desc = MID(MDL_desc,instr(MDL_desc,"|")+1,10000)
                If MDL_desc <> "" then MDL_desc_toel = MID(MDL_desc,1,instr(MDL_desc,"||")-2)
                else p=50 
             end if 
            
             'output "P = " & P
             'output "MDL_desc_versie= " & MDL_desc_versie
             'output "MDL_desc_basis= " & MDL_desc_basis
             'output "MDL_desc_datum= " & MDL_desc_datum
             'output "MDL_desc_auteur= "& MDL_desc_auteur
             'output "MDL_desc_taak= " & MDL_desc_taak
             'output "MDL_desc_toel= " & MDL_desc_toel
       
             set dummy = excel_set_cell_value(18+P, 1, "'" & MDL_desc_versie)
             set dummy = excel_set_cell_value(18+P, 2, "'" & MDL_desc_basis)
             set dummy = excel_set_cell_value(18+P, 3, "'" & MDL_desc_datum)
             set dummy = excel_set_cell_value(18+P, 4, "'" & MDL_desc_auteur)
             set dummy = excel_set_cell_value(18+P, 5, "'" & MDL_desc_taak)
             set dummy = excel_set_cell_value(18+P, 6, "'" & MDL_desc_toel)
             p=p+1
           Loop
      else 
        set dummy = excel_set_cell_value(19, 1, "Versie" )
        set dummy = excel_set_cell_value(19, 2, "Basisversie" )
        set dummy = excel_set_cell_value(19, 3, "Datum")
        set dummy = excel_set_cell_value(19, 4, "Auteur" )
        set dummy = excel_set_cell_value(19, 5, "Taak" )
        set dummy = excel_set_cell_value(19, 6, "Toelichting" )
      end if 
     
    set dummy = excel_freeze_row(3)
    set dummy = excel_autofit_columns(1, 6)
    set dummy = excel_set_cursor(1, 11)

    set dummy = excel_set_cell_value(100, 1, "'" & activemodel.author)
    set dummy = excel_set_cell_value(100, 2, "'" & activemodel.version)

   
    '------------------------------
 Dim changelog
     changelog = Split( activemodel.comment  , Chr(13))
      i = UBound(changelog)
      changelogteller = 101

   For j = 0 To i
           set dummy = excel_set_cell_value(changelogteller, 1 ,replace (changelog(j) , "=", "_") )     'begint met changelog(0), changelog(1)  etc  
           changelogteller = changelogteller + 1
    
   Next 'j
'------------------------------
    

 set excel_create_Titel_en_Historie = nothing  
end function 'excel_create_Titel_en_Historie

function excel_create_headers4LGM()      

pdm_filename = activemodel.filename                                                         
      do while instr(pdm_filename, "\") > 0                                                       
     pdm_filename = right(pdm_filename, len(pdm_filename) - instr(pdm_filename, "\"))          
loop   

bronsysteem =right(pdm_filename, 7 )                                                        
   bronsysteem =replace (bronsysteem , ".pdm", "" )                                            
   pdm_version = activemodel.version                                                     
   
       
 '--------------------------------------------------------------------------------------
'>>> sheet 2: entities
    set dummy = excel_select_worksheet(2)
     set dummy = excel_set_worksheet_name(2, "entities", true)     
     set dummy = excel_set_worksheet_tabcolor(2,  excel_colorindex_target   )   
    
    'set headers
     set dummy = excel_set_cell_value(1, 4, "SOURCE SCHEMA")                                                                              
     set dummy = excel_set_cell_value(1, 8, "TARGET SCHEMA")                                                                             
     set dummy = excel_set_cell_value(2, 1, "I")                                            'table.annotation            1 
     set dummy = excel_set_cell_value(2, 2, "U")                                            'table.annotation            2 
     set dummy = excel_set_cell_value(2, 3, "D")                                            'table.annotation            3 
     set dummy = excel_set_cell_value(2, 4, "USE-CASE")                                     'table.beginscript           4 
     set dummy = excel_set_cell_value(2, 5, "ENTITY USE-CASE")                              'table.name                  5 
     set dummy = excel_set_cell_value(2, 6, "FILE-NAME /ACRONIEM")                          'table.CheckConstraintName   6 
     set dummy = excel_set_cell_value(2, 7, "CALCULATION / TRANSFORMATION RULE")            'table.description           7 
     set dummy = excel_set_cell_value(2, 8, "ENVIRONMENT")                                  'table.endscript             8 
     set dummy = excel_set_cell_value(2, 9, "ENTITY")                                       'table.code                  9 
     set dummy = excel_set_cell_value(2, 10, "ENTITY STEREOTYPE")                           'table.stereotype            10
     set dummy = excel_set_cell_value(2, 11, "CALCULATION / TRANSFORMATION")                'table.type                  11
     set dummy = excel_set_cell_value(2, 12, "ENTITY REMARKS")                              'table.comment               12   
                                                                                           
        'set comments the two faces of the interface    
        'powerdesigner propertyset
         set dummy = excel_set_cell_comment (1, 1, "table.annotation")                       'I                                 1                  
         set dummy = excel_set_cell_comment (1, 2, "table.annotation")                       'U                                 2                  
         set dummy = excel_set_cell_comment (1, 3, "table.annotation")                       'D                                 3                  
         set dummy = excel_set_cell_comment (1, 4, "table.beginscript")                      'USE-CASE                          4                  
         set dummy = excel_set_cell_comment (1, 5, "table.name")                             'FILE                              5                  
         set dummy = excel_set_cell_comment (1, 6, "table.CheckConstraintName")              'CheckConstraintName               6                  
         set dummy = excel_set_cell_comment (1, 7, "table.description")                      'CALCULATION / TRANSFORMATION      7                  
         set dummy = excel_set_cell_comment (1, 8, "table.endscript")                        'ENVIRONMENT                       8                  
         set dummy = excel_set_cell_comment (1, 9, "table.code")                             'ENTITY                            9                  
         set dummy = excel_set_cell_comment (1, 10, "table.stereotype")                      'ENTITY STEREOTYPE                 10                 
         set dummy = excel_set_cell_comment (1, 11, "table.type")                            'CALCULATION / TRANSFORMATION      11                 
         set dummy = excel_set_cell_comment (1, 12, "table.comment")                         'ENTITY REMARKS                    12                 
      
        'source propertyset 
         set dummy = excel_set_cell_comment(2, 1, "I")                                       'table.annotation            1          
         set dummy = excel_set_cell_comment(2, 2, "U")                                       'table.annotation            2          
         set dummy = excel_set_cell_comment(2, 3, "D")                                       'table.annotation            3          
         set dummy = excel_set_cell_comment(2, 4, "USE-CASE")                                'table.beginscript           4          
         set dummy = excel_set_cell_comment(2, 5, "ENTITY USE-CASE")                         'table.name                  5          
         set dummy = excel_set_cell_comment(2, 6, "FILE-NAME /ACRONIEM")                     'table.CheckConstraintName   6          
         set dummy = excel_set_cell_comment(2, 7, "CALCULATION / TRANSFORMATION RULE")       'table.description           7          
         set dummy = excel_set_cell_comment(2, 8, "ENVIRONMENT")                             'table.endscript             8          
         set dummy = excel_set_cell_comment(2, 9, "ENTITY")                                  'table.code                  9          
         set dummy = excel_set_cell_comment(2, 10, "ENTITY STEREOTYPE")                      'table.stereotype            10         
         set dummy = excel_set_cell_comment(2, 11, "CALCULATION / TRANSFORMATION")           'table.type                  11         
         set dummy = excel_set_cell_comment(2, 12, "ENTITY REMARKS")                         'table.comment               12         
        
         'coloring  
         set dummy = excel_set_font_columns(1,12)                                                                                       
         set dummy = excel_set_cells_background_color(1, 1, 2, 7, excel_colorindex_source)                                              
         set dummy = excel_set_cells_font_color(1, 1, 2, 7, excel_colorindex_font_source)                                               
         set dummy = excel_set_cells_background_color(1, 8, 2, 12, excel_colorindex_target)                                            
         set dummy = excel_set_cells_font_color(1, 8, 2, 12, excel_colorindex_font_target)                                             
         set dummy = excel_set_cells_bold(1, 1, 1, 12, true)       
     
         'hiding   
        ' excel_obj.worksheets(3).columns(1).hidden= false ' IUD                                                                     
         
        'freeze
        set dummy = excel_freeze_row(3)           
        set dummy = excel_set_autofilter(2, 1, 2, 12)  ' excel_obj.range("A2:.Z2").AutoFilter         
         set dummy = excel_set_cursor(1, 1)                
        set dummy = excel_autofit_columns(1, 12)            

'--------------------------------------------------------------------------------------
  '>>> sheet 3: attributes
          set dummy = excel_select_worksheet(3)
          set dummy = excel_set_worksheet_name(3, "attributes", true)
          set dummy = excel_set_worksheet_tabcolor(3,  excel_colorindex_target   ) 

          'set headers
          set dummy = excel_set_cell_value(1, 4, "SOURCE SCHEMA")                       
          set dummy = excel_set_cell_value(1, 15, "TARGET SCHEMA")             
          set dummy = excel_set_cell_value(2, 1, "I")                                       'column.annotation          1          
          set dummy = excel_set_cell_value(2, 2, "U")                                       'column.annotation          2 
          set dummy = excel_set_cell_value(2, 3, "D")                                       'column.annotation          3 
          set dummy = excel_set_cell_value(2, 4, "ENTITY USE-CASE")                         'table.name                 4 
          set dummy = excel_set_cell_value(2, 5, "FIELD")                                   'column.name                5 
          set dummy = excel_set_cell_value(2, 6, "FIELD-DATATYPE")                          'column.format              6 
          set dummy = excel_set_cell_value(2, 7, "FIELD-SEQUENCE")                          'column.physicaloptions     7 
          set dummy = excel_set_cell_value(2, 8, "FIELD-STARTPOSITION")                     'column.lowvalue            8 
          set dummy = excel_set_cell_value(2, 9,  "FIELD-ENDPOSITION")                      'column.highvalue           9 
          set dummy = excel_set_cell_value(2, 10, "FIELD-LENGTH")                           'column.unit                10
          set dummy = excel_set_cell_value(2, 11, "PK/AK")                                  'column.primary             11
          set dummy = excel_set_cell_value(2, 12, "MANDATORY")                              'column.mandatory           12
          set dummy = excel_set_cell_value(2, 13, "PK NULLABLE")                            'column.nospace             13
          set dummy = excel_set_cell_value(2, 14, "CALCULATION / TRANSFORMATION")           'column.description         14
          set dummy = excel_set_cell_value(2, 15, "ENVIRONMENT")                            'table.endscript            15
          set dummy = excel_set_cell_value(2, 16, "ENTITY")                                 'table.code                 16
          set dummy = excel_set_cell_value(2, 17, "ATTRIBUTE")                              'column.code                17
          set dummy = excel_set_cell_value(2, 18, "DATA TYPE")                              'column.datatype            18
          set dummy = excel_set_cell_value(2, 19, "ATTRIBUTE STEREOTYPE")                   'column.stereotype          19
          set dummy = excel_set_cell_value(2, 20, "CALCULATION / TRANSFORMATION")           'column.ComputedExpression  20
          set dummy = excel_set_cell_value(2, 21, "PK/AK")                                  'column.primary             21
          set dummy = excel_set_cell_value(2, 22, "MANDATORY")                              'column.mandatory           22
          set dummy = excel_set_cell_value(2, 23, "PK NULLABLE")                            'column.nospace             23
          set dummy = excel_set_cell_value(2, 24, "ATTRIBUTE REMARKS")                      'column.comment             24
          
          'set comments the two faces of the interface    
          'powerdesigner propertyset
           set dummy = excel_set_cell_comment (1, 1 , "column.annotation         ")  '  1      I                           
           set dummy = excel_set_cell_comment (1, 2 , "column.annotation         ")  '  2      U                           
           set dummy = excel_set_cell_comment (1, 3 , "column.annotation         ")  '  3      D                           
           set dummy = excel_set_cell_comment (1, 4 , "table.name                ")  '  4      FILE                        
           set dummy = excel_set_cell_comment (1, 5 , "column.name               ")  '  5      FIELD                       
           set dummy = excel_set_cell_comment (1, 6 , "column.format             ")  '  6      FIELD-DATATYPE              
           set dummy = excel_set_cell_comment (1, 7 , "column.physicaloptions    ")  '  7      FIELD-SEQUENCE              
           set dummy = excel_set_cell_comment (1, 8 , "column.lowvalue           ")  '  8      FIELD-STARTPOSITION         
           set dummy = excel_set_cell_comment (1, 9 , "column.highvalue          ")  '  9      FIELD-ENDPOSITION           
           set dummy = excel_set_cell_comment (1, 10, "column.unit               ")  '  10     FIELD-LENGTH                
           set dummy = excel_set_cell_comment (1, 11, "column.primary            ")  '  11     PK/AK                       
           set dummy = excel_set_cell_comment (1, 12, "column.mandatory          ")  '  12     MANDATORY                   
           set dummy = excel_set_cell_comment (1, 13, "column.nospace            ")  '  13     PK NULLABLE                    
           set dummy = excel_set_cell_comment (1, 14, "column.description        ")  '  14     CALCULATION / TRANSFORMATION
           set dummy = excel_set_cell_comment (1, 15, "table.endscript           ")  '  15     ENVIRONMENT                 
           set dummy = excel_set_cell_comment (1, 16, "table.code                ")  '  16     ENTITY                      
           set dummy = excel_set_cell_comment (1, 17, "column.code               ")  '  17     ATTRIBUTE                   
           set dummy = excel_set_cell_comment (1, 18, "column.datatype           ")  '  18     DATA TYPE                   
           set dummy = excel_set_cell_comment (1, 19, "column.stereotype         ")  '  19     ATTRIBUTE STEREOTYPE        
           set dummy = excel_set_cell_comment (1, 20, "column.ComputedExpression ")  '  20     CALCULATION / TRANSFORMATION
           set dummy = excel_set_cell_comment (1, 21, "column.primary            ")  '  11     PK/AK                       
           set dummy = excel_set_cell_comment (1, 22, "column.mandatory          ")  '  12     MANDATORY                   
           set dummy = excel_set_cell_comment (1, 23, "column.nospace            ")  '  13     PK NULLABLE          
           set dummy = excel_set_cell_comment (1, 24, "column.comment            ")  '  21     ATTRIBUTE REMARKS                    
                                                                                                                                                              
          'propetyset source  
           set dummy = excel_set_cell_comment (2, 1 , "I                              ")  '  1   I                           
           set dummy = excel_set_cell_comment (2, 2 , "U                              ")  '  2   U                           
           set dummy = excel_set_cell_comment (2, 3 , "D                              ")  '  3   D                           
           set dummy = excel_set_cell_comment (2, 4 , "ENTITY USE-CASE                ")  '  4   FILE                        
           set dummy = excel_set_cell_comment (2, 5 , "FIELD                          ")  '  5   FIELD                       
           set dummy = excel_set_cell_comment (2, 6 , "FIELD-DATATYPE                 ")  '  6   FIELD-DATATYPE              
           set dummy = excel_set_cell_comment (2, 7 , "FIELD-SEQUENCE                 ")  '  7   FIELD-SEQUENCE              
           set dummy = excel_set_cell_comment (2, 8 , "FIELD-STARTPOSITION            ")  '  8   FIELD-STARTPOSITION         
           set dummy = excel_set_cell_comment (2, 9 , "FIELD-ENDPOSITION              ")  '  9   FIELD-ENDPOSITION           
           set dummy = excel_set_cell_comment (2, 10, "FIELD-LENGTH                   ")  '  10  FIELD-LENGTH                
           set dummy = excel_set_cell_comment (2, 11, "PK/AK                          ")  '  11  PK/AK                       
           set dummy = excel_set_cell_comment (2, 12, "MANDATORY                      ")  '  12  MANDATORY                   
           set dummy = excel_set_cell_comment (2, 13, "PK NULLABLE                    ")  '  13  NULLABLE                    
           set dummy = excel_set_cell_comment (2, 14, "CALCULATION / TRANSFORMATION   ")  '  14  CALCULATION / TRANSFORMATION
           set dummy = excel_set_cell_comment (2, 15, "ENVIRONMENT                    ")  '  15  ENVIRONMENT                 
           set dummy = excel_set_cell_comment (2, 16, "ENTITY                         ")  '  16  ENTITY                      
           set dummy = excel_set_cell_comment (2, 17, "ATTRIBUTE                      ")  '  17  ATTRIBUTE                   
           set dummy = excel_set_cell_comment (2, 18, "DATA TYPE                      ")  '  18  DATA TYPE                   
           set dummy = excel_set_cell_comment (2, 19, "ATTRIBUTE STEREOTYPE           ")  '  20  ATTRIBUTE STEREOTYPE        
           set dummy = excel_set_cell_comment (2, 20, "CALCULATION / TRANSFORMATION   ")  '  19  CALCULATION / TRANSFORMATION
           set dummy = excel_set_cell_comment (2, 21, "PK/AK                          ")  '  20  PK/AK                       
           set dummy = excel_set_cell_comment (2, 22, "MANDATORY                      ")  '  21  MANDATORY                   
           set dummy = excel_set_cell_comment (2, 23, "PK NULLABLE                    ")  '  22  NULLABLE       
           set dummy = excel_set_cell_comment (2, 24, "ATTRIBUTE REMARKS              ")  '  23  ATTRIBUTE REMARKS           
                                                                                                                                                              
          'coloring                                                                                                                                                 
           set dummy = excel_set_cells_background_color(1, 1, 1, 14, excel_colorindex_source)                                                                       
           set dummy = excel_set_cells_font_color(1, 1, 2, 14, excel_colorindex_font_source)                                                                        
           set dummy = excel_set_cells_background_color(1, 15, 2, 24, excel_colorindex_target)                                                                      
           set dummy = excel_set_cells_font_color(1, 15, 2, 24, excel_colorindex_font_target)                                                                       
           set dummy = excel_set_cells_bold(1, 1, 1, 24, true)                                                                                                      
                                                                                                                                                                    
           set dummy = excel_set_cells_background_color(2, 1,  2,  1,  excel_colorindex_source)    'I                                  column.annotation	1          
           set dummy = excel_set_cells_background_color(2, 2,  2,  2,  excel_colorindex_source)    'U                                  column.CannotModify 	2      
           set dummy = excel_set_cells_background_color(2, 3,  2,  3,  excel_colorindex_source)    'D                                  column.uppercase	3          
           set dummy = excel_set_cells_background_color(2, 4,  2,  4,  excel_colorindex_source)    'FILE                               table.name	4                
           set dummy = excel_set_cells_background_color(2, 5,  2,  5,  excel_colorindex_source)    'FIELD                              column.name	5                
           set dummy = excel_set_cells_background_color(2, 6,  2,  6,  excel_colorindex_source)    'FIELD-DATATYPE                     column.format	6              
           set dummy = excel_set_cells_background_color(2, 7,  2,  7,  excel_colorindex_source)    'FIELD-SEQUENCE                     column.physicaloptions	7    
           set dummy = excel_set_cells_background_color(2, 8,  2,  8,  excel_colorindex_source)    'FIELD-STARTPOSITION                column.lowvalue 	8          
           set dummy = excel_set_cells_background_color(2, 9,  2,  9,  excel_colorindex_source)    'FIELD-ENDPOSITION                  column.highvalue	9          
           set dummy = excel_set_cells_background_color(2, 10, 2, 10,  excel_colorindex_source)    'FIELD-LENGTH                       column.unit	10               
           set dummy = excel_set_cells_background_color(2, 11 ,2, 11, excel_colorindex_source)     'PK/AK                              column.primary	11           
           set dummy = excel_set_cells_background_color(2, 12, 2, 12, excel_colorindex_source)     'MANDATORY                          column.mandatory	12         
           set dummy = excel_set_cells_background_color(2, 13, 2, 13, excel_colorindex_source)     'NOSPACE                           column.nospace	13           
           set dummy = excel_set_cells_background_color(2, 14, 2, 14, excel_colorindex_source)     'CALCULATION / TRANSFORMATION       column.description	14       
           set dummy = excel_set_cells_background_color(2, 15, 2, 15, excel_colorindex_target)     'ENVIRONMENT                        table.endscript	15           
           set dummy = excel_set_cells_background_color(2, 16, 2, 16, excel_colorindex_target)     'ENTITY                             table.code	16               
           set dummy = excel_set_cells_background_color(2, 17, 2, 17, excel_colorindex_target)     'ATTRIBUTE                          column.code	17               
           set dummy = excel_set_cells_background_color(2, 18, 2, 18, excel_colorindex_target)     'DATA TYPE                          column.datatype	18           
           set dummy = excel_set_cells_background_color(2, 19, 2, 19, excel_colorindex_target)     'ATTRIBUTE STEREOTYPE               column.PhysicalOptions	19   
           set dummy = excel_set_cells_background_color(2, 20, 2, 20, excel_colorindex_target)     'CALCULATION / TRANSFORMATION       column.stereotype	20     
           set dummy = excel_set_cells_background_color(2, 21 ,2, 21, excel_colorindex_source)     'PK/AK                              column.primary	21           
           set dummy = excel_set_cells_background_color(2, 22, 2, 22, excel_colorindex_source)     'MANDATORY                          column.mandatory	22         
           set dummy = excel_set_cells_background_color(2, 23, 2, 23, excel_colorindex_source)     'NOSPACE                           column.nospace	23           
           set dummy = excel_set_cells_background_color(2, 24, 2, 24, excel_colorindex_target)     'ATTRIBUTE REMARKS                  column.comment	24      
             
          'coloring  
          set dummy = excel_set_font_columns(1,21)                                                                                       
          set dummy = excel_set_cells_background_color(1, 1, 1, 14, excel_colorindex_source)                                              
          set dummy = excel_set_cells_font_color(1, 1, 2, 14, excel_colorindex_font_source)                                               
          set dummy = excel_set_cells_background_color(1,15, 2, 24, excel_colorindex_target)                                            
          set dummy = excel_set_cells_font_color(1, 15, 2, 24, excel_colorindex_font_target)                                             
          set dummy = excel_set_cells_bold(1, 1, 1, 24, true)                                                                            
          
           'hiding   
          ' excel_obj.worksheets(3).columns(1).hidden= false ' IUD
          
         'freeze
         set dummy = excel_freeze_row(3)                             
         set dummy = excel_set_autofilter(2, 1, 2, 24)  ' excel_obj.range("A2:.Z2").AutoFilter
         set dummy = excel_set_cursor(1, 1) 
         set dummy = excel_autofit_columns(1, 24)            
         
         
   
 '--------------------------------------------------------------------------------------
 '>>> sheet 4: relations
       set dummy = excel_select_worksheet(4)
       set dummy = excel_set_worksheet_name(4, "entity relationships", true)
       set dummy = excel_set_worksheet_tabcolor(4, excel_colorindex_target  ) 
    
        'set headers      
        set dummy = excel_set_cell_value(1, 4, "SOURCE SCHEMA")  
        set dummy = excel_set_cell_value(1, 14, "TARGET SCHEMA") 
       
        set dummy = excel_set_cell_value(2,1   , "I"                                   )  'reference.annotation               1
        set dummy = excel_set_cell_value(2,2   , "U"                                   )  'reference.annotation               2 
        set dummy = excel_set_cell_value(2,3   , "D"                                   )  'reference.annotation               3 
        set dummy = excel_set_cell_value(2,4   , "reference.name"                      )  'reference.name                     4 
        set dummy = excel_set_cell_value(2,5   , "reference.childtable.name"           )  'reference.childtable.name          5 
        set dummy = excel_set_cell_value(2,6   , "reference.parenttable.name"          )  'reference.parenttable.name         6 
        set dummy = excel_set_cell_value(2,7   , "reference.cardinality"               )  'reference.cardinality              7 
        set dummy = excel_set_cell_value(2,8   , "reference.DeleteConstraint"          )  'reference.DeleteConstraint         8 
        set dummy = excel_set_cell_value(2,9   , "reference.UpdateConstraint"          )  'reference.UpdateConstraint         9 
        set dummy = excel_set_cell_value(2,10  , "reference.ChangeParentAllowed"       )  'reference.ChangeParentAllowed      10   
        set dummy = excel_set_cell_value(2,11  , "reference.childrole"                 )  'reference.childrole                11
        set dummy = excel_set_cell_value(2,12  , "reference.parentrole"                )  'reference.parentrole               12
        set dummy = excel_set_cell_value(2,13  , "reference.comment"                   )  'reference.comment                  13
        set dummy = excel_set_cell_value(2,14  , "reference.code"                      )  'reference.code                     14
        set dummy = excel_set_cell_value(2,15  , "reference.childtable.code"           )  'reference.childtable.code          15
        set dummy = excel_set_cell_value(2,16  , "reference.parenttable.code"          )  'reference.parenttable.code         16
        set dummy = excel_set_cell_value(2,17  , "reference.stereotype"                )  'reference.stereotype               17
        set dummy = excel_set_cell_value(2,18  , "reference.ImplementationType"        )  'reference.ImplementationType       18
        set dummy = excel_set_cell_value(2,19  , "reference.description"               )  'reference.description              19
        set dummy = excel_set_cell_value(2,20  , "reference.JoinExpression"            )  'reference.JoinExpression           20
        set dummy = excel_set_cell_value(2,21  , "reference.ForeignKeyColumnList"      )  'reference.ForeignKeyColumnList     21
        set dummy = excel_set_cell_value(2,22  , "reference.ParentKeyColumnList"       )  'reference.ParentKeyColumnList      22
        set dummy = excel_set_cell_value(2,23  , "reference.foreignkeyconstraintname"  )  'reference.foreignkeyconstraintname 23
                                                                                          
       'set comments the two faces of the interface    
          'powerdesigner propertyset
           set dummy = excel_set_cell_comment (1, 1 ,"I")                                   ' reference.annotation                 1 
           set dummy = excel_set_cell_comment (1, 2 ,"U")                                   ' reference.annotation                 2 
           set dummy = excel_set_cell_comment (1, 3 ,"D")                                   ' reference.annotation                 3 
           set dummy = excel_set_cell_comment (1, 4 ,"reference.name")                      ' reference.name                       4 
           set dummy = excel_set_cell_comment (1, 5 ,"reference.childtable.name")           ' reference.childtable.name            5 
           set dummy = excel_set_cell_comment (1, 6 ,"reference.parenttable.name")          ' reference.parenttable.name           6 
           set dummy = excel_set_cell_comment (1, 7 ,"reference.cardinality")               ' reference.cardinality                7 
           set dummy = excel_set_cell_comment (1, 8 ,"reference.DeleteConstraint")          ' reference.DeleteConstraint           8 
           set dummy = excel_set_cell_comment (1, 9 ,"reference.UpdateConstraint")          ' reference.UpdateConstraint           9 
           set dummy = excel_set_cell_comment (1, 10,"reference.ChangeParentAllowed")       ' reference.ChangeParentAllowed        10
           set dummy = excel_set_cell_comment (1, 11,"reference.childrole")                 ' reference.childrole                  11
           set dummy = excel_set_cell_comment (1, 12,"reference.parentrole")                ' reference.parentrole                 12
           set dummy = excel_set_cell_comment (1, 13,"reference.comment")                   ' reference.comment                    13
           set dummy = excel_set_cell_comment (1, 14,"reference.code")                      ' reference.code                       14
           set dummy = excel_set_cell_comment (1, 15,"reference.childtable.code")           ' reference.childtable.code            15
           set dummy = excel_set_cell_comment (1, 16,"reference.parenttable.code")          ' reference.parenttable.code           16
           set dummy = excel_set_cell_comment (1, 17,"reference.stereotype")                ' reference.stereotype                 17
           set dummy = excel_set_cell_comment (1, 18,"LNK met SAT")                         ' reference.ImplementationType         18
           set dummy = excel_set_cell_comment (1, 19, "reference.description" )             ' reference.description                19
           set dummy = excel_set_cell_comment (1, 20,"reference.JoinExpression")            ' reference.JoinExpression             20
           set dummy = excel_set_cell_comment (1, 21,"reference.ForeignKeyColumnList")      ' reference.ForeignKeyColumnList       21  
           set dummy = excel_set_cell_comment (1, 22,"reference.ParentKeyColumnList")       ' reference.ParentKeyColumnList        22
           set dummy = excel_set_cell_comment (1, 23,"reference.foreignkeyconstraintname")  ' reference.foreignkeyconstraintname   23

         'source propertyset 
          
        'coloring  
        set dummy = excel_set_font_columns(1,12)                                                                                       
        set dummy = excel_set_cells_background_color(1, 1, 2, 13, excel_colorindex_source)                                              
        set dummy = excel_set_cells_font_color(1, 1, 2, 13, excel_colorindex_font_source)                                               
        set dummy = excel_set_cells_background_color(1, 14, 2, 23, excel_colorindex_target)                                            
        set dummy = excel_set_cells_font_color(1, 14, 2, 23, excel_colorindex_font_target)                                             
        set dummy = excel_set_cells_bold(1, 1, 1, 23, true)                                                                            
       
       'freeze
       set dummy = excel_freeze_row(3)                             
       set dummy = excel_set_autofilter(2, 1, 2, 23)  ' excel_obj.range("A2:.Z2").AutoFilter
       set dummy = excel_set_cursor(1, 1) 
       set dummy = excel_autofit_columns(1, 23)    
       
    
 '--------------------------------------------------------------------------------------
 '>>> sheet 5: joins
       set dummy = excel_select_worksheet(5)
       set dummy = excel_set_worksheet_name(5, "ERD association_attributes" , true)
       set dummy = excel_set_worksheet_tabcolor(5,  excel_colorindex_target   )   
         'set headers      
          set dummy = excel_set_cell_value(1, 1, "SOURCE SCHEMA")                                  
          set dummy = excel_set_cell_value(1, 6, "TARGET SCHEMA")                                 
          set dummy = excel_set_cell_value(2, 1   , "reference.name                    "    )    '  reference.name                       1      
          set dummy = excel_set_cell_value(2, 2   , "reference.childtable.name         "    )    '  reference.childtable.name            2      
          set dummy = excel_set_cell_value(2, 3   , "join.childtablecolumn.name        "    )    '  join.childtablecolumn.name           3      
          set dummy = excel_set_cell_value(2, 4   , "reference.parenttable.name        "    )    '  reference.parenttable.name           4      
          set dummy = excel_set_cell_value(2, 5   , "join.parenttablecolumn.name       "    )    '  join.parenttablecolumn.name          5      
          set dummy = excel_set_cell_value(2, 6   , "reference.code                    "    )    '  reference.code                       6      
          set dummy = excel_set_cell_value(2, 7   , "reference.childtable.code         "    )    '  reference.childtable.code            7      
          set dummy = excel_set_cell_value(2, 8   , "join.childtablecolumn.code        "    )    '  join.childtablecolumn.code           8      
          set dummy = excel_set_cell_value(2, 9   , "reference.parenttable.code        "    )    '  reference.parenttable.code           9   
          set dummy = excel_set_cell_value(2, 10  , "join.parenttablecolumn.code       "    )    '  join.parenttablecolumn.code          10

         'set comments the two faces of the interface    
        'powerdesigner propertyset
         set dummy = excel_set_cell_comment (2, 1  , "reference.name                   ")      ' reference.name                       1       
         set dummy = excel_set_cell_comment (2, 2  , "reference.childtable.name        ")      ' reference.childtable.name            2       
         set dummy = excel_set_cell_comment (2, 3  , "join.childtablecolumn.name       ")      ' join.childtablecolumn.name           3       
         set dummy = excel_set_cell_comment (2, 4  , "reference.parenttable.name       ")      ' reference.parenttable.name           4       
         set dummy = excel_set_cell_comment (2, 5  , "join.parenttablecolumn.name      ")      ' join.parenttablecolumn.name          5       
         set dummy = excel_set_cell_comment (2, 6  , "reference.code                   ")      ' reference.code                       6       
         set dummy = excel_set_cell_comment (2, 7  , "reference.childtable.code        ")      ' reference.childtable.code            7       
         set dummy = excel_set_cell_comment (2, 8  , "join.childtablecolumn.code       ")      ' join.childtablecolumn.code           8       
         set dummy = excel_set_cell_comment (2, 9  , "reference.parenttable.code       ")      ' reference.parenttable.code           9       
         set dummy = excel_set_cell_comment (2, 10 , "join.parenttablecolumn.code      ")      ' join.parenttablecolumn.code          10      

         'coloring  
         set dummy = excel_set_font_columns(1,12)                                                                                       
         set dummy = excel_set_cells_background_color(1, 1, 2, 5, excel_colorindex_source)                                              
         set dummy = excel_set_cells_font_color(1, 1, 2, 5, excel_colorindex_font_source)                                               
         set dummy = excel_set_cells_background_color(1, 6, 2, 10, excel_colorindex_target)                                            
         set dummy = excel_set_cells_font_color(1, 6, 2, 10, excel_colorindex_font_target)                                             
         set dummy = excel_set_cells_bold(1, 1, 1, 10, true)                               
                                                                                                                                       
       'freeze
        set dummy = excel_freeze_row(3)                             
        set dummy = excel_set_autofilter(2, 1, 2, 10)  ' excel_obj.range("A2:.Z2").AutoFilter
        set dummy = excel_set_cursor(1, 1) 
        set dummy = excel_autofit_columns(1, 10)   
        
        
 '--------------------------------------------------------------------------------------
 '>>> sheet 6: views
        set dummy = excel_select_worksheet(6)
        set dummy = excel_set_worksheet_name(6, "views" , true )
        set dummy = excel_set_worksheet_tabcolor(6,  excel_colorindex_magenta   )
        
        set dummy = excel_set_cell_value(1, 4, pdm_version)         
        set dummy = excel_set_cell_value(1, 5, pdm_filename)       
        'set headers
        
        set dummy = excel_set_cell_value(2, 1  ,  "I")                     '    view.annotation                 1
        set dummy = excel_set_cell_value(2, 2  ,  "U")                     '    view.annotation                 2
        set dummy = excel_set_cell_value(2, 3  ,  "D")                     '    view.annotation                 3
        set dummy = excel_set_cell_value(2, 4  ,  "view.code")             '    view.code                       4
        set dummy = excel_set_cell_value(2, 5  ,  "view.name")             '    view.name                       5
        set dummy = excel_set_cell_value(2, 6  ,  "view.stereotype")       '    view.stereotype                 6
        set dummy = excel_set_cell_value(2, 7  ,  "view.comment")          '    view.comment                    7
        set dummy = excel_set_cell_value(2, 8  ,  "view.type")             '    view.type                       8
        set dummy = excel_set_cell_value(2, 9  ,  "view.description")      '    view.description                9
        set dummy = excel_set_cell_value(2, 10  ,  "view.sqlquery")        '    view.sqlquery                   10
        
        'set comments the two faces of the interface
        set dummy = excel_set_cell_comment (2,  1,     "view.annotation ")  
        set dummy = excel_set_cell_comment (2,  2,     "view.annotation ")  
        set dummy = excel_set_cell_comment (2,  3,     "view.annotation ")  
        set dummy = excel_set_cell_comment (2,  4,     "view.code")  
        set dummy = excel_set_cell_comment (2,  5,     "view.name")  
        set dummy = excel_set_cell_comment (2,  6,     "view.stereotype")  
        set dummy = excel_set_cell_comment (2,  7,     "view.comment")  
        set dummy = excel_set_cell_comment (2,  8,     "view.type")  
        set dummy = excel_set_cell_comment (2,  9,     "view.description")  
        set dummy = excel_set_cell_comment (2, 10,     "view.sqlquery")  

       'set comments the two faces of the interface    
        'coloring  
        set dummy = excel_set_font_columns(1,12)                                                                                       
        set dummy = excel_set_cells_background_color(1, 1, 1, 10, excel_colorindex_source)                                              
        set dummy = excel_set_cells_font_color(1, 1, 2, 8, excel_colorindex_font_source)                                               
        set dummy = excel_set_cells_background_color(1, 1, 2, 10, excel_colorindex_target)                                            
        set dummy = excel_set_cells_font_color(1, 1, 2, 10, excel_colorindex_font_target)                                             
        set dummy = excel_set_cells_bold(1, 1, 1, 10, true)                                                                            
     
       'freeze
       set dummy = excel_freeze_row(3)                             
       set dummy = excel_set_autofilter(2, 1, 2, 10)  ' excel_obj.range("A2:.Z2").AutoFilter 
       set dummy = excel_set_cursor(1, 1) 
       set dummy = excel_autofit_columns(1, 10)            
      
       
    
'--------------------------------------------------------------------------------------
   '>>> sheet 7: viewcolumns
         set dummy = excel_select_worksheet(7)
        
        set dummy = excel_set_worksheet_name(7, "viewcolumns" , true )
        set dummy = excel_set_worksheet_tabcolor(7,  excel_colorindex_magenta   ) 
        
         set dummy = excel_set_cell_value(1, 4, pdm_version)         
         set dummy = excel_set_cell_value(1, 5, pdm_filename)       
       
        'set headers
         set dummy = excel_set_cell_value(2, 1   ,  "I")                          '    column.annotation               1 
         set dummy = excel_set_cell_value(2, 2   ,  "U")                          '    column.annotation               2 
         set dummy = excel_set_cell_value(2, 3   ,  "D")                          '    column.annotation               3 
         set dummy = excel_set_cell_value(2, 4   ,  "view.code")                  '    view.code                       4 
         set dummy = excel_set_cell_value(2, 5   ,  "column.code")                '    column.code                     5 
         set dummy = excel_set_cell_value(2, 6   ,  "column.name")                '    column.name                     6 
         set dummy = excel_set_cell_value(2, 7   ,  "column.format")              '    column.format                   7 
         set dummy = excel_set_cell_value(2, 8   ,  "len.")                       '    column.unit                     8 
         set dummy = excel_set_cell_value(2, 9   ,  "pos_van.")                   '    column.lowvalue                 9 
         set dummy = excel_set_cell_value(2, 10  ,  "pos_tot")                    '    column.highvalue                10
         set dummy = excel_set_cell_value(2, 11  ,  "column.datatype")            '    column.datatype                 11
         set dummy = excel_set_cell_value(2, 12  ,  "column.stereotype")          '    column.stereotype               12
         set dummy = excel_set_cell_value(2, 13  ,  "column.comment")             '    column.comment                  13
         set dummy = excel_set_cell_value(2, 14  ,  "column.description")         '    column.description              14
      
       'set comments the two faces of the interface                         
         set dummy = excel_set_cell_comment (2,  1 ,  "column.annotation      ")      '   I                            1 
         set dummy = excel_set_cell_comment (2,  2 ,  "column.annotation      ")      '   U                            2 
         set dummy = excel_set_cell_comment (2,  3 ,  "column.annotation      ")      '   D                            3 
         set dummy = excel_set_cell_comment (2,  4 ,  "view.code              ")      '   view.code                    4 
         set dummy = excel_set_cell_comment (2,  5 ,  "column.code            ")      '   column.code                  5 
         set dummy = excel_set_cell_comment (2,  6 ,  "column.name            ")      '   column.name                  6 
         set dummy = excel_set_cell_comment (2,  7 ,  "column.format          ")      '   column.format                7 
         set dummy = excel_set_cell_comment (2,  8 ,  "column.unit            ")      '   len.                         8 
         set dummy = excel_set_cell_comment (2,  9 ,  "column.lowvalue        ")      '   pos_van.                     9 
         set dummy = excel_set_cell_comment (2,  10,  "column.highvalue       ")      '   pos_tot                      10
         set dummy = excel_set_cell_comment (2,  11,  "column.datatype        ")      '   column.datatype              11
         set dummy = excel_set_cell_comment (2,  12,  "column.stereotype      ")      '   column.stereotype            12
         set dummy = excel_set_cell_comment (2,  13,  "column.comment         ")      '   column.comment               13
         set dummy = excel_set_cell_comment (2,  14,  "column.description     ")      '   column.description           14
         
       'coloring  
       'set dummy = excel_set_cells_background_color(1, 1, 1, 9, excel_colorindex_source)                                              
       'set dummy = excel_set_cells_font_color(1, 1, 2, 8, excel_colorindex_font_source)                                               
       set dummy = excel_set_cells_background_color(1, 1, 2, 14, excel_colorindex_target)                                            
       set dummy = excel_set_cells_font_color(1, 1, 2, 14, excel_colorindex_font_target)                                             
       set dummy = excel_set_cells_bold(1, 1, 1, 14, true)                                                                            
        
       'freeze
       set dummy = excel_freeze_row(3)                             
       set dummy = excel_set_autofilter(2, 1, 2, 14)  ' excel_obj.range("A2:.Z2").AutoFilter 
       set dummy = excel_set_cursor(1, 1) 
       set dummy = excel_autofit_columns(1, 14)     
       
    
 '--------------------------------------------------------------------------------------
   '>>> sheet 8: diagrams
         set dummy = excel_select_worksheet(8)
         set dummy = excel_set_worksheet_name(8, "diagrams" , true)
         set dummy = excel_set_worksheet_tabcolor(8,  excel_colorindex_gray   ) 
         
         set dummy = excel_set_cell_value(1, 1, pdm_version)     
         set dummy = excel_set_cell_value(1, 2, pdm_filename)    
         'set headers
         ' set dummy = excel_set_cell_value(1, 1, "SOURCE SCHEMA")                                                                                      
         ' set dummy = excel_set_cell_value(1, 6, "TARGET SCHEMA")                                                                                      
          set dummy = excel_set_cell_value(2, 1   , "diagram.code"    )             '  diagram.code                              1                     
          set dummy = excel_set_cell_value(2, 2   , "diagram.name"    )             '  diagram.name                              2                     
          set dummy = excel_set_cell_value(2, 3   , "diagram.pageformat"    )       '  diagram.pageformat                        3    
          set dummy = excel_set_cell_value(2, 4   , "diagram.pageorientation"    )  '  diagram.pageorientation                   4                                      
         
         'set comments the two faces of the interface    
          'powerdesigner propertyset
            set dummy = excel_set_cell_comment (1, 1, "diagram.code")
            set dummy = excel_set_cell_comment (1, 2, "diagram.name")
            set dummy = excel_set_cell_comment (1, 3, "diagram.pageformat")
            set dummy = excel_set_cell_comment (1, 4, "diagram.pageorientation")
           
          'source propertyset 
           set dummy = excel_set_cell_comment(2, 1, "VERSION") 
             
          'coloring  
        '  set dummy = excel_set_font_columns(1,12)                                                                                       
       ' set dummy = excel_set_cells_background_color(1, 1, 1, 8, excel_colorindex_source)                                              
       ' set dummy = excel_set_cells_font_color(1, 1, 2, 8, excel_colorindex_font_source)                                               
        set dummy = excel_set_cells_background_color(1, 1, 2, 4, excel_colorindex_target)                                            
        set dummy = excel_set_cells_font_color(1, 1, 2, 4, excel_colorindex_font_target)                                             
        set dummy = excel_set_cells_bold(1, 1, 1, 4, true)                                                                            
        
     'freeze
      set dummy = excel_freeze_row(3)                             
    
     set dummy = excel_set_autofilter(2, 1, 2, 4)  ' excel_obj.range("A2:.Z2").AutoFilter
     set dummy = excel_set_cursor(1, 1) 
     set dummy = excel_autofit_columns(1, 4)            
     
    
    '--------------------------------------------------------------------------------------
  '>>> sheet 9: symbols
         set dummy = excel_select_worksheet(9)
         set dummy = excel_set_worksheet_name(9, "symbols" , true )
         set dummy = excel_set_worksheet_tabcolor(9,  excel_colorindex_gray   ) 
         
         set dummy = excel_set_cell_value(1, 1, pdm_version)         
         set dummy = excel_set_cell_value(1, 2, pdm_filename)        
     
        'set headers
         set dummy = excel_set_cell_value(2, 1    ,  "diagram.code            ")       '    diagram.code                   1 
         set dummy = excel_set_cell_value(2, 2    ,  "symbol.objecttype       ")       '    symbol.objecttype              2 
         set dummy = excel_set_cell_value(2, 3    ,  "symbol.code             ")       '    symbol.code                    3 
         set dummy = excel_set_cell_value(2, 4    ,  "symbol.name             ")       '    symbol.name                    4 
         set dummy = excel_set_cell_value(2, 5    ,  "symbol.rect.top         ")       '    symbol.rect.top                5 
         set dummy = excel_set_cell_value(2, 6    ,  "symbol.rect.bottom      ")       '    symbol.rect.bottom             6 
         set dummy = excel_set_cell_value(2, 7    ,  "symbol.rect.left        ")       '    symbol.rect.left               7 
         set dummy = excel_set_cell_value(2, 8    ,  "symbol.rect.right       ")       '    symbol.rect.right              8 
         set dummy = excel_set_cell_value(2, 9    ,  "symbol.linecolor        ")       '    symbol.linecolor               9 
         set dummy = excel_set_cell_value(2, 10   ,  "symbol.fillcolor        ")       '    symbol.fillcolor               10
         set dummy = excel_set_cell_value(2, 11   ,  "symbol.gradientfillmode ")       '    symbol.gradientfillmode        11
         set dummy = excel_set_cell_value(2, 12   ,  "symbol.AutoAdjustToText ")       '    symbol.AutoAdjustToText        12
         set dummy = excel_set_cell_value(2, 13   ,  "symbol.KeepAspect       ")       '    symbol.KeepAspect              13
         set dummy = excel_set_cell_value(2, 14   ,  "symbol.KeepCenter       ")       '    symbol.KeepCenter              14
         set dummy = excel_set_cell_value(2, 15   ,  "symbol.KeepSize         ")       '    symbol.KeepSize                15
     '
     '    'set comments the two faces of the interface    
     '     'powerdesigner propertyset
     '      set dummy = excel_set_cell_comment (1, 1, "table.type") 
     '      
     '     'source propertyset 
     '      set dummy = excel_set_cell_comment(2, 1, "VERSION") 
     '        
          'coloring  
     '     set dummy = excel_set_cells_background_color(1, 1, 1, 8, excel_colorindex_source)                                              
     '     set dummy = excel_set_cells_font_color(1, 1, 2, 8, excel_colorindex_font_source)                                               

        set dummy = excel_set_cells_background_color(1, 1, 2, 15, excel_colorindex_target)                                            
        set dummy = excel_set_cells_font_color(1, 1, 2, 15, excel_colorindex_font_target)                                             
        set dummy = excel_set_cells_bold(1, 1, 1, 15, true)                                                                            
     
       'freeze
        set dummy = excel_freeze_row(3)                             
        set dummy = excel_autofit_columns(1, 15)            
        set dummy = excel_set_cursor(1, 1) 
        set dummy = excel_set_autofilter(2, 1, 2, 15)  ' excel_obj.range("A2:.Z2").AutoFilter
        
        
        '10 
        '11
        '12
        '13
        '14
        '15
        '16
        
           '--------------------------------------------------------------------------------------
     '>>> sheet 17: viewreferences  
        set dummy = excel_select_worksheet(17)
        set dummy = excel_set_worksheet_name(17, "view references", true)
        set dummy = excel_set_worksheet_tabcolor(17,  excel_colorindex_magenta   ) 
        
        set dummy = excel_set_cell_value(2, 1  , "I"                              )'  view reference.annotation                1 
        set dummy = excel_set_cell_value(2, 2  , "U"                              )'  view reference.annotation                2 
        set dummy = excel_set_cell_value(2, 3  , "D"                              )'  view reference.annotation                3 
        set dummy = excel_set_cell_value(2, 4  , "vwref.name"                     )'  view reference.name                      4 
        set dummy = excel_set_cell_value(2, 5  , "vwref.childtable.name"          )'  view reference.childtable.name           5 
        set dummy = excel_set_cell_value(2, 6  , "vwref.parenttable.name"         )'  view reference.parenttable.name          6 
        set dummy = excel_set_cell_value(2, 7  , "vwref.cardinality"              )'  view reference.cardinality               7 
        set dummy = excel_set_cell_value(2, 8  , "vwref.DeleteConstraint"         )'  view reference.DeleteConstraint          8 
        set dummy = excel_set_cell_value(2, 9  , "vwref.UpdateConstraint"         )'  view reference.UpdateConstraint          9 
        set dummy = excel_set_cell_value(2,10  , "vwref.ChangeParentAllowed"      )'  view reference.ChangeParentAllowed       10
        set dummy = excel_set_cell_value(2,11  , "vwref.childrole"                )'  view reference.childrole                 11
        set dummy = excel_set_cell_value(2,12  , "vwref.parentrole"               )'  view reference.parentrole                12
        set dummy = excel_set_cell_value(2,13  , "vwref.comment"                  )'  view reference.comment                   13
        set dummy = excel_set_cell_value(2,14  , "vvwref.code"                    )'  view reference.code                      14
        set dummy = excel_set_cell_value(2,15  , "vwref.childtable.code"          )'  view reference.childtable.code           15
        set dummy = excel_set_cell_value(2,16  , "vwref.parenttable.code"         )'  view reference.parenttable.code          16
        set dummy = excel_set_cell_value(2,17  , "vwref.stereotype"               )'  view reference.stereotype                17
        set dummy = excel_set_cell_value(2,18  , "vwref.ImplementationType"       )'  view reference.ImplementationType        18
        set dummy = excel_set_cell_value(2,19  , "vwref.description"              )'  view reference.description               19
        set dummy = excel_set_cell_value(2,20  , "vwref.JoinExpression"           )'  view reference.JoinExpression            20
        set dummy = excel_set_cell_value(2,21  , "vwref.ForeignKeyColumnList"     )'  view reference.ForeignKeyColumnList      21
        set dummy = excel_set_cell_value(2,22  , "vwref.ParentKeyColumnList"      )'  view reference.ParentKeyColumnList       22
        set dummy = excel_set_cell_value(2,23  , "vwref.foreignkeyconstraintname" )'  view reference.foreignkeyconstraintname  23
                                                                                      
         
        'coloring  
        set dummy = excel_set_cells_background_color(1, 1, 2, 13, excel_colorindex_target)                                              
        set dummy = excel_set_cells_font_color(1, 1, 2, 13, excel_colorindex_font_target)                                                       
        set dummy = excel_set_cells_background_color(1, 14, 2, 23, excel_colorindex_target)                                            
        set dummy = excel_set_cells_font_color(1, 14, 2, 23, excel_colorindex_font_target)                                             
               
       'freeze
       set dummy = excel_freeze_row(3)                             
       
       set dummy = excel_set_autofilter(2, 1, 2, 23)  ' excel_obj.range("A2:.Z2").AutoFilter
       set dummy = excel_set_cell_value(1, 4, STI_modelnaam)
       set dummy = excel_set_cell_value(1, 5, pdm_version)
       set dummy = excel_set_cell_value(1, 6, CDW_modelnaam)
       set dummy = excel_set_cursor(1, 1) 
       set dummy = excel_autofit_columns(1, 23)            
       
        
    
 '--------------------------------------------------------------------------------------
 '>>> sheet 18: viewreference joins
       set dummy = excel_select_worksheet(18)
       set dummy = excel_set_worksheet_name(18 ,"viewref joins" , true)
       set dummy = excel_set_worksheet_tabcolor(18,  excel_colorindex_magenta   ) 
                                 
          set dummy = excel_set_cell_value(2, 1   , "vwref.name                        "    )    '  reference.name                       1      
          set dummy = excel_set_cell_value(2, 2   , "vwref childtable.view             "    )    '  reference.childtable.name            2      
          set dummy = excel_set_cell_value(2, 3   , "vwrefjoin.childviewcolumn.name    "    )    '  join.childtablecolumn.name           3      
          set dummy = excel_set_cell_value(2, 4   , "vwref parentview.name             "    )    '  reference.parenttable.name           4      
          set dummy = excel_set_cell_value(2, 5   , "vwrefjoin.parenttablecolumn.name  "    )    '  join.parenttablecolumn.name          5      
          set dummy = excel_set_cell_value(2, 6   , "vwref.code                        "    )    '  reference.code                       6      
          set dummy = excel_set_cell_value(2, 7   , "vwref.childview.code              "    )    '  reference.childtable.code            7      
          set dummy = excel_set_cell_value(2, 8   , "vwrefjoin.childtview.olumn.code   "    )    '  join.childtablecolumn.code           8      
          set dummy = excel_set_cell_value(2, 9   , "vwref.parentviewcode              "    )    '  reference.parenttable.code           9   
          set dummy = excel_set_cell_value(2, 10  , "vwrefjoin.parentview.column.code  "    )    '  join.parenttablecolumn.code          10

         'set comments the two faces of the interface    
        'powerdesigner propertyset
         set dummy = excel_set_cell_comment (2, 1  , "vwref.name                       "    )      ' reference.name                       1       
         set dummy = excel_set_cell_comment (2, 2  , "vwref childtable.view            "    )      ' reference.childtable.name            2       
         set dummy = excel_set_cell_comment (2, 3  , "vwrefjoin.childviewcolumn.name   "    )      ' join.childtablecolumn.name           3       
         set dummy = excel_set_cell_comment (2, 4  , "vwref parentview.name            "    )      ' reference.parenttable.name           4       
         set dummy = excel_set_cell_comment (2, 5  , "vwrefjoin.parenttablecolumn.name "    )      ' join.parenttablecolumn.name          5       
         set dummy = excel_set_cell_comment (2, 6  , "vwref.code                       "    )      ' reference.code                       6       
         set dummy = excel_set_cell_comment (2, 7  , "vwref.childview.code             "    )      ' reference.childtable.code            7       
         set dummy = excel_set_cell_comment (2, 8  , "vwrefjoin.childtview.olumn.code  "    )      ' join.childtablecolumn.code           8       
         set dummy = excel_set_cell_comment (2, 9  , "vwref.parentviewcode             "    )      ' reference.parenttable.code           9       
         set dummy = excel_set_cell_comment (2, 10 , "vwrefjoin.parentview.column.code "    )      ' join.parenttablecolumn.code          10      
                                                                                            
         'coloring  
         set dummy = excel_set_cells_background_color(1, 1, 2, 5, excel_colorindex_target)                                              
         set dummy = excel_set_cells_font_color(1, 1, 2, 5, excel_colorindex_font_target)                                               
         
         set dummy = excel_set_cells_background_color(1, 6, 2, 10, excel_colorindex_target)                                            
         set dummy = excel_set_cells_font_color(1, 6, 2, 10, excel_colorindex_font_target)                                             
                                                                                                                                       
       'freeze
        set dummy = excel_freeze_row(3)                             
       
        set dummy = excel_set_autofilter(2, 1, 2, 10)  ' excel_obj.range("A2:.Z2").AutoFilter
        set dummy = excel_set_cell_value(1, 1, STI_modelnaam)
        set dummy = excel_set_cell_value(1, 2, pdm_version)
        set dummy = excel_set_cell_value(1, 3, CDW_modelnaam)
        set dummy = excel_set_cursor(1, 1) 
        set dummy = excel_autofit_columns(1, 10)            
        
        
                                             
 set excel_create_headers4LGM = nothing      
end function 'excel_create_headers4LGM       


'=====================================
function excel_hide_worksheets ()

excel_obj.worksheets(1).visible = true  'titel en versiebeheer
excel_obj.worksheets(2).visible = true  'entities
excel_obj.worksheets(3).visible = true  'attributes
excel_obj.worksheets(4).visible = true  'relations
excel_obj.worksheets(5).visible = true  'joins
excel_obj.worksheets(6).visible = false  'views
excel_obj.worksheets(7).visible = false  'view attributes  
excel_obj.worksheets(8).visible = false  'diagrams
excel_obj.worksheets(9).visible = false  'symbols

excel_obj.worksheets(10).visible = true ' 
excel_obj.worksheets(11).visible = true ' 
excel_obj.worksheets(12).visible = true ' 
excel_obj.worksheets(13).visible = true ' 
excel_obj.worksheets(14).visible = true ' 
excel_obj.worksheets(15).visible = true '
excel_obj.worksheets(16).visible = true '
excel_obj.worksheets(17).visible = false 'view references
excel_obj.worksheets(18).visible = false 'view references joins

set excel_hide_worksheets = nothing 

end function 'excel_hide_worksheets  


'=====================================
function excel_hide_columns()

'------------------------entities
excel_obj.worksheets(2).columns(1 ).hidden=false     ' table_annotation                 table.annotation                  
excel_obj.worksheets(2).columns(2 ).hidden=false     ' table_annotation                 table.annotation                   
excel_obj.worksheets(2).columns(3 ).hidden=false     ' table_annotation                 table.annotation                   
excel_obj.worksheets(2).columns(4 ).hidden=false     ' table_beginscript                table.beginscript                  
excel_obj.worksheets(2).columns(5 ).hidden=false     ' table_name                       table.name                         
excel_obj.worksheets(2).columns(6 ).hidden=false     ' table_CheckConstraintName        table.CheckConstraintName         acroniem
excel_obj.worksheets(2).columns(7 ).hidden=false     ' table_description                table.description                  
excel_obj.worksheets(2).columns(8 ).hidden=false     ' table_endscript                  table.endscript                    
excel_obj.worksheets(2).columns(9 ).hidden=false     ' table_code                       table.code                         
excel_obj.worksheets(2).columns(10).hidden=false     ' table_stereotype                 table.stereotype                   
excel_obj.worksheets(2).columns(11).hidden=false     ' table_type                       table.type                         
excel_obj.worksheets(2).columns(12).hidden=false     ' table_comment                    table.comment                      
'------------------------attributes
excel_obj.worksheets(3).columns(1 ).hidden=false     ' column_annotation                column.annotation                  
excel_obj.worksheets(3).columns(2 ).hidden=false     ' column_annotation                column.annotation                  
excel_obj.worksheets(3).columns(3 ).hidden=false     ' column_annotation                column.annotation                  
excel_obj.worksheets(3).columns(4 ).hidden=false     ' table_name                       table.name                         
excel_obj.worksheets(3).columns(5 ).hidden=false     ' column_name                      column.name                        
excel_obj.worksheets(3).columns(6 ).hidden=false     ' column_format                    column.format                      
excel_obj.worksheets(3).columns(7 ).hidden=false     ' column_physicaloptions           column.physicaloptions             
excel_obj.worksheets(3).columns(8 ).hidden=false     ' column_lowvalue                  column.lowvalue                    
excel_obj.worksheets(3).columns(9 ).hidden=false     ' column_highvalue                 column.highvalue                   
excel_obj.worksheets(3).columns(10).hidden=false     ' column_unit                      column.unit                        
excel_obj.worksheets(3).columns(11).hidden=false     ' column_primary                   column.primary                     
excel_obj.worksheets(3).columns(12).hidden=false     ' column_mandatory                 column.mandatory                   
excel_obj.worksheets(3).columns(13).hidden=false     ' column_nospace                   column.nospace                     
excel_obj.worksheets(3).columns(14).hidden=false     ' column_description               column.description                 
excel_obj.worksheets(3).columns(15).hidden=false     ' table_endscript                  table.endscript                    
excel_obj.worksheets(3).columns(16).hidden=false     ' table_code                       table.code                         
excel_obj.worksheets(3).columns(17).hidden=false     ' column_code                      column.code                        
excel_obj.worksheets(3).columns(18).hidden=false     ' column_datatype                  column.datatype                    
excel_obj.worksheets(3).columns(19).hidden=false     ' column_stereotype                column.stereotype                  
excel_obj.worksheets(3).columns(20).hidden=false     ' column_ComputedExpression        column.ComputedExpression          
excel_obj.worksheets(3).columns(21).hidden=false     ' column_comment                   column.comment                     
'------------------------relations
excel_obj.worksheets(4).columns(1 ).hidden=false     ' reference_annotation             reference.annotation               
excel_obj.worksheets(4).columns(2 ).hidden=false     ' reference_annotation             reference.annotation               
excel_obj.worksheets(4).columns(3 ).hidden=false     ' reference_annotation             reference.annotation               
excel_obj.worksheets(4).columns(4 ).hidden=false     ' reference_name                   reference.name                     
excel_obj.worksheets(4).columns(5 ).hidden=false     ' reference_childtable_name        reference.childtable.name          
excel_obj.worksheets(4).columns(6 ).hidden=false     ' reference_parenttable_name       reference.parenttable.name         
excel_obj.worksheets(4).columns(7 ).hidden=false     ' reference_cardinality            reference.cardinality              
excel_obj.worksheets(4).columns(8 ).hidden=false     ' reference_DeleteConstraint       reference.DeleteConstraint         
excel_obj.worksheets(4).columns(9 ).hidden=false     ' reference_UpdateConstraint       reference.UpdateConstraint         
excel_obj.worksheets(4).columns(10).hidden=false     ' reference_ChangeParentAllowed    reference.ChangeParentAllowed      
excel_obj.worksheets(4).columns(11).hidden=false     ' reference_childrole              reference.childrole                
excel_obj.worksheets(4).columns(12).hidden=false     ' reference_parentrole             reference.parentrole               
excel_obj.worksheets(4).columns(13).hidden=false     ' reference_comment                reference.comment                  
excel_obj.worksheets(4).columns(14).hidden=false     ' reference_code                   reference.code                     
excel_obj.worksheets(4).columns(15).hidden=false     ' reference_childtable_code        reference.childtable.code          
excel_obj.worksheets(4).columns(16).hidden=false     ' reference_parenttable_code       reference.parenttable.code         
excel_obj.worksheets(4).columns(17).hidden=false     ' reference_stereotype             reference.stereotype               
excel_obj.worksheets(4).columns(18).hidden=false     ' reference_ImplementationType     reference.ImplementationType       
excel_obj.worksheets(4).columns(19).hidden=false     ' reference_description            reference.description              
excel_obj.worksheets(4).columns(20).hidden=false     ' reference_JoinExpression         reference.JoinExpression           
excel_obj.worksheets(4).columns(21).hidden=false     ' reference_ForeignKeyColumnList   reference.ForeignKeyColumnList     
excel_obj.worksheets(4).columns(22).hidden=false     ' reference_ParentKeyColumnList    reference.ParentKeyColumnList      
excel_obj.worksheets(4).columns(23).hidden=false     ' reference_foreignkeyconstraintnamreference.foreignkeyconstraintnamee
'------------------------joins
excel_obj.worksheets(5).columns(1 ).hidden=false     ' reference_name                   reference.name                     
excel_obj.worksheets(5).columns(2 ).hidden=false     ' reference_childtable_name        reference.childtable.name          
excel_obj.worksheets(5).columns(3 ).hidden=false     ' join_childtablecolumn_name       join.childtablecolumn.name         
excel_obj.worksheets(5).columns(4 ).hidden=false     ' reference_parenttable_name       reference.parenttable.name         
excel_obj.worksheets(5).columns(5 ).hidden=false     ' join_parenttablecolumn_name      join.parenttablecolumn.name        
excel_obj.worksheets(5).columns(6 ).hidden=false     ' reference_code                   reference.code                     
excel_obj.worksheets(5).columns(7 ).hidden=false     ' reference_childtable_code        reference.childtable.code          
excel_obj.worksheets(5).columns(8 ).hidden=false     ' join_childtablecolumn_code       join.childtablecolumn.code         
excel_obj.worksheets(5).columns(9 ).hidden=false     ' reference_parenttable_code       reference.parenttable.code         
excel_obj.worksheets(5).columns(10).hidden=false     ' join_parenttablecolumn_code      join.parenttablecolumn.code        
'------------------------views                          
excel_obj.worksheets(6).columns(1 ).hidden=false     ' view_annotation                  view.annotation                    
excel_obj.worksheets(6).columns(2 ).hidden=false     ' view_annotation                  view.annotation                    
excel_obj.worksheets(6).columns(3 ).hidden=false     ' view_annotation                  view.annotation                    
excel_obj.worksheets(6).columns(4 ).hidden=false     ' view_code                        view.code                          
excel_obj.worksheets(6).columns(5 ).hidden=false     ' view_name                        view.name                          
excel_obj.worksheets(6).columns(6 ).hidden=false     ' view_stereotype                  view.stereotype                    
excel_obj.worksheets(6).columns(7 ).hidden=false     ' view_comment                     view.comment                       
excel_obj.worksheets(6).columns(8 ).hidden=false     ' view_type                        view.type                          
excel_obj.worksheets(6).columns(9 ).hidden=false     ' view_description                 view.description                   
excel_obj.worksheets(6).columns(10).hidden=false     ' view_sqlquery                    view.sqlquery                      
'------------------------viewscolumns                       
excel_obj.worksheets(7).columns(1 ).hidden=false     ' column_annotation                column.annotation                  
excel_obj.worksheets(7).columns(2 ).hidden=false     ' column_annotation                column.annotation                  
excel_obj.worksheets(7).columns(3 ).hidden=false     ' column_annotation                column.annotation                  
excel_obj.worksheets(7).columns(4 ).hidden=false     ' view_code                        view.code                          
excel_obj.worksheets(7).columns(5 ).hidden=false     ' column_code                      column.code                        
excel_obj.worksheets(7).columns(6 ).hidden=false     ' column_name                      column.name                        
excel_obj.worksheets(7).columns(7 ).hidden=false     ' column_format                    column.format                      
excel_obj.worksheets(7).columns(8 ).hidden=false     ' column_unit                      column.unit                        
excel_obj.worksheets(7).columns(9 ).hidden=false     ' column_lowvalue                  column.lowvalue                    
excel_obj.worksheets(7).columns(10).hidden=false     ' column_highvalue                 column.highvalue                   
excel_obj.worksheets(7).columns(11).hidden=false     ' column_mandatory                 column.mandatory                   
excel_obj.worksheets(7).columns(12).hidden=false     ' column_stereotype                column.stereotype                  
excel_obj.worksheets(7).columns(13).hidden=false     ' column_comment                   column.comment                     
excel_obj.worksheets(7).columns(14).hidden=false     ' column_description               column.description                 
'------------------------diagrams
excel_obj.worksheets(8).columns(1 ).hidden=false     ' diagram_code                     diagram.code                       
excel_obj.worksheets(8).columns(2 ).hidden=false     ' diagram_name                     diagram.name                       
excel_obj.worksheets(8).columns(3 ).hidden=false     ' diagram_pageformat               diagram.pageformat                 
excel_obj.worksheets(8).columns(4 ).hidden=false     ' diagram_pageorientation          diagram.pageorientation            
'------------------------symbols
excel_obj.worksheets(9).columns(1 ).hidden=false     ' diagram_code                     diagram.code                       
excel_obj.worksheets(9).columns(2 ).hidden=false     ' symbol_objecttype                symbol.objecttype                  
excel_obj.worksheets(9).columns(3 ).hidden=false     ' symbol_code                      symbol.code                        
excel_obj.worksheets(9).columns(4 ).hidden=false     ' symbol_name                      symbol.name                        
excel_obj.worksheets(9).columns(5 ).hidden=false     ' symbol_rect_top                  symbol.rect.top                    
excel_obj.worksheets(9).columns(6 ).hidden=false     ' symbol_rect_bottom               symbol.rect.bottom                 
excel_obj.worksheets(9).columns(7 ).hidden=false     ' symbol_rect_left                 symbol.rect.left                   
excel_obj.worksheets(9).columns(8 ).hidden=false     ' symbol_rect_right                symbol.rect.right                  
excel_obj.worksheets(9).columns(9 ).hidden=false     ' symbol_linecolor                 symbol.linecolor                   
excel_obj.worksheets(9).columns(10).hidden=false     ' symbol_fillcolor                 symbol.fillcolor                   
excel_obj.worksheets(9).columns(11).hidden=false     ' symbol_gradientfillmode          symbol.gradientfillmode            
excel_obj.worksheets(9).columns(12).hidden=false     ' symbol_AutoAdjustToText          symbol.AutoAdjustToText            
excel_obj.worksheets(9).columns(13).hidden=false     ' symbol_KeepAspect                symbol.KeepAspect                  
excel_obj.worksheets(9).columns(14).hidden=false     ' symbol_KeepCenter                symbol.KeepCenter                  
excel_obj.worksheets(9).columns(15).hidden=false     ' symbol_KeepSize                  symbol.KeepSize                    


set excel_hide_columns = nothing 
end function 'excel_hide_columns




'*******************************************************************   
 '>>> sheet 10:           
 function excel_create_sheet10()                                             
         
  '>>> sheet 10 '"start (Ontwerp)
      set dummy = excel_select_worksheet(10) 
     
      set dummy = excel_set_worksheet_name(10, "start (Ontwerp)", true)
      set dummy = excel_set_cell_value(1, 1, pdm_version)
      set dummy = excel_set_cell_value(1, 2, pdm_filename)
      
      set dummy = excel_set_cell_value(2, 1, "Interface beschrijving")
      set dummy = excel_set_cell_value(2, 2, "Create Datawarehouse Extract")
      set dummy = excel_set_cell_value(2, 3, " Status: "& pdm_version )
      set dummy = excel_set_cell_value(2, 4, date & " - " & time )
      
      set dummy = excel_set_cells_bold(1, 1, 1, 2, true)
      set dummy = excel_set_cells_background_color(1, 1, 1, 4, excel_colorindex_avro)
      set dummy = excel_set_cells_background_color(2, 1, 2, 4, excel_colorindex_header)
      set dummy = excel_set_cells_font_color(1, 1, 2, 4, excel_colorindex_font_header)
      
      set dummy = excel_set_cell_value(6, 1, "")
      set dummy = excel_set_cell_value(6, 2, "Voor een bronontsluiting hebben we de volgende gegevens nodig:")
      set dummy = excel_set_cells_bold(6, 1, 6, 1, true)
      set dummy = excel_set_cells_background_color(6, 1, 6, 2, excel_colorindex_gray)
      set dummy = excel_set_cells_border_color(6, 1, 6, 2, excel_colorindex_black) 
     
      set dummy = excel_set_cell_value(7, 1, "[1]")
      set dummy = excel_set_cell_value(7, 2, "entiteiten")
      set dummy = excel_set_cells_border_color(7, 1, 7, 2, excel_colorindex_black) 
      set dummy = excel_set_cell_value(8, 1, "[2]")
      set dummy = excel_set_cell_value(8, 2, "van iedere entiteit de attributen")
      set dummy = excel_set_cells_border_color(8, 1, 8, 2, excel_colorindex_black) 
      set dummy = excel_set_cell_value(9, 1, "[3]")
      set dummy = excel_set_cell_value(9, 2, "van ieder attribuut het datatype (òf, als het datatype niet beschikbaar is, het maximaal aantal karakters waarmee het attribuut gevuld kan worden)")
      set dummy = excel_set_cells_border_color(9, 1, 9, 2, excel_colorindex_black) 
      set dummy = excel_set_cell_value(10, 1, "[4]")
      set dummy = excel_set_cell_value(10, 2, "de primaire sleutel van iedere entiteit (dus welke attributen maken een enteitrij uniek)")
      set dummy = excel_set_cells_border_color(10, 1, 10, 2, excel_colorindex_black) 
      set dummy = excel_set_cell_value(11, 1, "[5]")
      set dummy = excel_set_cell_value(11, 2, "de relaties tussen de entiteiten (als deze niet worden aangeleverd of niet kunnen worden gegarandeerd, dan nemen we ze niet op in ons model)")
      set dummy = excel_set_cells_border_color(11, 1, 11, 2, excel_colorindex_black) 
      set dummy = excel_set_cell_value(12, 1, "[6]")
      set dummy = excel_set_cell_value(12, 2, "de associerende attributen die onderdeel zijn van de relaties.")
      set dummy = excel_set_cells_border_color(12, 1, 12, 2, excel_colorindex_black) 
     
      set dummy = excel_set_font_columns(1, 17) 
      excel_obj.range("A1:Z2").font.name = "Arial"'behalve de eerste twee regels
     
      set dummy = excel_freeze_row(3)
      set dummy = excel_autofit_columns(1, 4)
      set dummy = excel_set_cursor(1, 1) 
      set dummy = excel_set_autofilter(2, 1, 2, 4)  ' excel_obj.range("A2:.Z2").AutoFilter
       
  set excel_create_sheet10 = nothing      
end function 'excel_create_sheet10                                      
                                                                       
                                                                      
'*******************************************************************   
 '>>> sheet 11:                                                        
 function excel_create_sheet11    ()                                         
         
  '>>> sheet 11 "datatypes (Bouw)

     set dummy = excel_select_worksheet(11) 
     set dummy = excel_set_worksheet_name(11, "datatypes (Bouw) ", true)
     set dummy = excel_set_cell_value(1, 1, pdm_version)
     set dummy = excel_set_cell_value(1, 2, pdm_filename)
 
     set dummy = excel_set_cell_value(2, 1, "Interface beschrijving")
     set dummy = excel_set_cell_value(2, 2, "Create Datawarehouse Extract")
     set dummy = excel_set_cell_value(2, 3, " Status: "& pdm_version )
     set dummy = excel_set_cell_value(2, 4, date & " - " & time )
     
     set dummy = excel_set_cells_bold(1, 1, 1, 2, true)
     set dummy = excel_set_cells_background_color(1, 1, 1, 4, excel_colorindex_avro)
     set dummy = excel_set_cells_background_color(2, 1, 2, 4, excel_colorindex_header)
     set dummy = excel_set_cells_font_color(1, 1, 2, 4, excel_colorindex_font_header)
     
     set dummy = excel_freeze_row(3)
     
     set dummy = excel_autofit_columns(1, 4)
     set dummy = excel_set_cursor(1, 1) 
     set dummy = excel_set_autofilter(2, 1, 2, 4)  ' excel_obj.range("A2:.Z2").AutoFilter
    
  
  set excel_create_sheet11 = nothing      
end function 'excel_create_sheet11       
                                                                       
'*******************************************************************   
                                                                       
'*******************************************************************   
 '>>> sheet 12:                                                        
  function excel_create_sheet12    ()                                         

'>>> sheet 12: special Cases
 'naming        
    set dummy = excel_select_worksheet(12)'
    set dummy = excel_set_worksheet_name(12, "special Cases", true)

'    set dummy = excel_set_cell_value(2, 1, "Belastingdienst")
'    set dummy = excel_set_cell_value(2, 2, "Data-requirements R2013.3 - EDW:")
'    set dummy = excel_set_cell_value(2, 3, "Intern afstemmings-document B/CA & B/CAO:" )
'    set dummy = excel_set_cell_value(2, 4, date & " - " & time )
    
    set dummy = excel_set_cells_bold(1, 1, 1, 2, true)
    set dummy = excel_set_cells_background_color(1, 1, 1, 4, excel_colorindex_avro)
    set dummy = excel_set_cells_background_color(2, 1, 2, 4, excel_colorindex_header)
    set dummy = excel_set_cells_font_color(1, 1, 2, 4, excel_colorindex_font_header)
     
     
    set dummy = excel_set_cell_value(3,  2, "Use this section to specify special cases and transformations	 " )  
    set dummy = excel_set_cells_font_color(3, 2, 3, 2, excel_colorindex_red)
    set dummy = excel_set_cells_bold(3, 2, 3, 2, true)

    set dummy = excel_set_cell_value(4,  2, "Special Case	 " )  
    set dummy = excel_set_cell_value(4,  3, "SPECxx " )  
    set dummy = excel_set_cell_value(5,  2, "Target table" )  
    set dummy = excel_set_cell_value(6,  2, "Target field" )  
    set dummy = excel_set_cell_value(7,  2, "Motivation/reason" )  
    set dummy = excel_set_cells_background_color(4, 2, 7, 2, excel_colorindex_gray) 
    set dummy = excel_set_cells_border_color(4, 2, 7, 2, excel_colorindex_black) 
    
    set dummy = excel_set_cell_value(8,  1, "VERSION" )  
    set dummy = excel_set_cell_value(8,  2, "Description" )  
    set dummy = excel_set_cells_background_color(8, 1, 8, 2, excel_colorindex_yellow) 
    set dummy = excel_set_cells_border_color(8, 1, 8, 2, excel_colorindex_black)    
    
    set dummy = excel_set_cell_value(9,  1, "0.1 <tbd>" )  

     set dummy = excel_set_cell_value(11,  2, "Special Case	 " )  
    set dummy = excel_set_cell_value(11,  3, "SPECyy" )  
    set dummy = excel_set_cell_value(12,  2, "Target table" )  
    set dummy = excel_set_cell_value(13,  2, "Target field" )  
    set dummy = excel_set_cell_value(14,  2, "Motivation/reason" )  
    set dummy = excel_set_cells_background_color(11, 2, 14, 2, excel_colorindex_gray) 
    set dummy = excel_set_cells_border_color(11, 2, 14, 2, excel_colorindex_black) 
    
    set dummy = excel_set_cell_value(15,  1, "VERSION" )  
    set dummy = excel_set_cell_value(15,  2, "Description" )  
    set dummy = excel_set_cells_background_color(15, 1, 15, 2, excel_colorindex_yellow) 
    set dummy = excel_set_cells_border_color(15, 1, 15, 2, excel_colorindex_black)    
    
    set dummy = excel_set_cell_value(16,  1, "0.1 <tbd>" )  

    'freeze
     set dummy = excel_freeze_row(3)                             
     set dummy = excel_autofit_columns(1, 12)            
     set dummy = excel_set_cursor(1, 1) 
     
  
  set excel_create_sheet12 = nothing      
end function 'excel_create_sheet12
                                                                       
'*******************************************************************   
 '>>> sheet 13:                                                        
 function excel_create_sheet13                                            
         
  'eventhook
    
  set excel_create_sheet13 = nothing      
end function 'excel_create_sheet13
                                                                       
'*******************************************************************   
                                                                       
                                                                       
'*******************************************************************   
 '>>> sheet 14:                                                        
function excel_create_sheet14                                            
         
  'eventhook
    
  set excel_create_sheet14 = nothing      
end function 'excel_create_sheet14
                                                                       
'*******************************************************************   
 '>>> sheet 15:                                                        
 function excel_create_sheet15                                            
         
  'eventhook
    
  set excel_create_sheet15 = nothing      
end function 'excel_create_sheet15
                                                                       
'*******************************************************************   
 '>>> sheet 16:                                                        
 function excel_create_sheet16                                            
         
  'eventhook
    
  set excel_create_sheet16 = nothing      
end function 'excel_create_sheet14
                                                                       
'*******************************************************************   
 
'********** matrix functions **********
 
'for storing data, we use a matrix with a number of rows and a number of columns given by nr_matrix_rows and nr_matrix_columns, respectively
'the matrix rows are numbered 0, 1, 2, ..., nr_matrix_rows - 1
'the matrix columns are numbered 0, 1, 2, ..., nr_matrix_columns - 1
 
'the matrix has three hidden columns that are used for storing layout information
'column nr_matrix_columns + 0 stores a value that is used for sorting the rows in ascending order
'column nr_matrix_columns + 1 stores a colorindex that is used for displaying the row
'column nr_matrix_columns + 2 stores a boolean that indicates whether or not the row must be displayed underlined
 

 
function matrix_sort_ascending(matrix, nr_matrix_rows, nr_matrix_columns)
 dim temp()
 redim temp(nr_matrix_columns + 3)
 for i = 0 to nr_matrix_rows - 2
for j = i + 1 to nr_matrix_rows - 1
 if matrix(i, nr_matrix_columns + 0) > matrix(j, nr_matrix_columns + 0) then
for matrix_column_nr = 0 to nr_matrix_columns + 2
 temp(matrix_column_nr)= matrix(i, matrix_column_nr)
 matrix(i, matrix_column_nr) = matrix(j, matrix_column_nr)
 matrix(j, matrix_column_nr) = temp(matrix_column_nr)
next
 end if
next
 next
 set matrix_sort_ascending = nothing
end function
 
function matrix_write_to_excel(matrix, nr_matrix_rows, nr_matrix_columns, excel_start_row_nr, excel_start_column_nr)
 for matrix_row_nr = 0 to nr_matrix_rows - 1
if matrix_row_nr = 0 then excel_row_nr = excel_start_row_nr else excel_row_nr = excel_row_nr + 1
if (par_insert_empty_lines and matrix_row_nr > 0) then
 if matrix(matrix_row_nr - 1, nr_matrix_columns + 1) <> matrix(matrix_row_nr, nr_matrix_columns + 1) then excel_row_nr = excel_row_nr + 1
end if
for matrix_column_nr = 0 to nr_matrix_columns - 1
 excel_column_nr = excel_start_column_nr + matrix_column_nr
 set dummy = excel_set_cell_value(excel_row_nr, excel_column_nr, matrix(matrix_row_nr, matrix_column_nr))
next
if par_alternating_colors then
 set dummy = excel_set_cells_font_color(excel_row_nr, excel_start_column_nr, excel_row_nr, excel_start_column_nr + nr_matrix_columns - 1, matrix(matrix_row_nr, nr_matrix_columns + 1))
else
 set dummy = excel_set_cells_font_color(excel_row_nr, excel_start_column_nr, excel_row_nr, excel_start_column_nr + nr_matrix_columns - 1, excel_colorindex_1)
end if
if par_underline_primary_keys then
 set dummy = excel_set_cells_underline(excel_row_nr, excel_start_column_nr, excel_row_nr, excel_start_column_nr + nr_matrix_columns - 1, matrix(matrix_row_nr, nr_matrix_columns + 2))
end if
 next
 set matrix_write_to_excel = nothing
end function
 
'Voor een bronontsluiting hebben we de volgende gegevens nodig:
'[1] entiteiten.
'[2] van iedere entiteit de attributen.
'[3] van ieder attribuut het datatype (òf, als het datatype niet beschikbaar is, het maximaal aantal karakters waarmee het attribuut gevuld kan worden).
'[4] de primaire sleutel van iedere entiteit (dus welke attributen maken een enteitrij uniek).
'[5] de relaties tussen de entiteiten (als deze niet worden aangeleverd of niet kunnen worden gegarandeerd, dan nemen we ze niet op in ons model).
'[6] de associerende attributen die onderdeel zijn van de relaties.
