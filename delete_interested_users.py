import configparser
import json
import requests
import sys
import time
import xlrd
import xlwt
import xlutils.copy


def main(input):
    st = time.localtime()
    start_time = time.strftime("%H:%M:%S", st)
    
    # Read config file
    config = configparser.ConfigParser()
    config.read('local_settings.ini')
    key = config['Alma Acq R/W']['key']
    
    # Read spreadsheet
    book_in = xlrd.open_workbook(input)
    sheet1 = book_in.sheet_by_index(0)
    pol_id_col_index = 0
    getpol_col_index = 1
    updatepol_col_index = 2
    
    # Copy spreadsheet for output
    book_out = xlutils.copy.copy(book_in)
    
    # Add new column headers
    book_out.get_sheet(0).write(0,getpol_col_index,'Get_POL')
    book_out.get_sheet(0).write(0,updatepol_col_index,'Update_POL')
    
    for row in range(1, sheet1.nrows):
        pol_id = sheet1.cell(row, pol_id_col_index).value
        print(row)
        
        #Get POL
        headers = {'accept':'application/json'}
        response = requests.get('https://api-na.hosted.exlibrisgroup.com/almaws/v1/acq/po-lines/'+pol_id+'?apikey='+key, headers=headers)
        book_out.get_sheet(0).write(row,getpol_col_index,response.status_code)
        print('Get POL: ', response.status_code)
        time.sleep(0.04)
        
        #Remove interested users and update POL
        if response.status_code == 200:
            pol_json = response.json()
            pol_json['interested_user'] = [{}]
            headers = {'accept':'application/json', 'Content-Type':'application/json'}
            response = requests.put('https://api-na.hosted.exlibrisgroup.com/almaws/v1/acq/po-lines/'+pol_id+'?update_inventory=false&redistribute_funds=false&apikey='+key, headers=headers, data=json.dumps(pol_json))
            book_out.get_sheet(0).write(row,updatepol_col_index,response.status_code)
            print('Update POL: ', response.status_code)
            time.sleep(0.04)
        print()
        book_out.save('results.xls')
    
    et = time.localtime()
    end_time = time.strftime("%H:%M:%S", et)
    print('Start Time: ', start_time)
    print('End Time: ', end_time)    


if __name__ == '__main__':
    main(sys.argv[1])
