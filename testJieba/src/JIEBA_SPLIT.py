#-*- coding: UTF-8 -*-
import jieba
import jieba.posseg as pseg  
import xlrd
import xlwt
import News_Object



def split(xls_name):
    jieba.set_dictionary("dict.txt")
    stopkey=[line.strip().decode('utf-8') for line in open("stop_word.txt").readlines()] 
    #開關excel===============================
    data=xlrd.open_workbook('UN_SPLIT/'+xls_name)
    table = data.sheet_by_name(u'TEST')
    nrows_num = table.nrows
    ncols_num = table.ncols
    res=[]
    wb = xlwt.Workbook()
    ws = wb.add_sheet('A Test Sheet')
    i=0
    temp_date =""
    temp_ur =""
    temp_title=""
    temp_content=""
    result=""
    print("START SPLIT___"+xls_name+".........") 
#=======================================
    for nrows in range(nrows_num):
        for ncols in range(ncols_num):
            if ncols ==0:
                temp_date = table.cell(nrows,ncols).value 
            elif ncols ==1:
                temp_ur = table.cell(nrows,ncols).value
            elif ncols ==2:
                temp_title = table.cell(nrows,ncols).value 
                
            else:
                cell_value = table.cell(nrows,ncols).value.encode('utf-8') 
                temp_content = table.cell(nrows,ncols).value
                words = pseg.cut(cell_value) #进行分词  
                  #记录最终结果的变量  
                
                for ex1 in words:
                    cut_word = ex1.word
                    cut_flag = ex1.flag
                    comp1 = ex1.word.encode('utf-8')
                    check_stop= False;
                   
                    for ex in stopkey:
                        comp = ex.encode('utf-8')              
                        if comp.decode('utf-8') == comp1.decode('utf-8'): 
                            check_stop = True;   
                    
                    if check_stop == False:
                        result += cut_word + ":" + cut_flag + ","            
            
        ws.write(i, 0, temp_date)   
        ws.write(i, 1, temp_ur)   
        ws.write(i, 2, temp_title)   
        ws.write(i, 3, temp_content)   
        ws.write(i, 4, result)   
            
        temp_date =""
        temp_ur =""
        temp_title=""
        temp_content=""    
        result=""
        i=i+1 
        print(i) 
                
            
    
    wb.save('SPLIT/'+'split_'+xls_name)
    print("Please check result file:SPLIT/"+xls_name) 
#=======================================
