/*
author:lidingyi
*/
use umya_spreadsheet::*;
use std::path::Path;




fn main(){
    let file_path = Path::new("./test.xlsx"); //文件目录
    let header_number = 1u32;
    let book = reader::xlsx::read(file_path).unwrap(); //读取文件
    let sheet = book.get_sheet_collection().first().unwrap(); //获取第一个sheet
    let headers = sheet.get_collection_by_row(&header_number);
    
    //循环每一行
    for i in sheet.get_row_dimensions(){
        if *i.get_row_num()>header_number{
            let mut book = new_file(); //创建新的文件
            let sheet_new = book.new_sheet("Sheet").unwrap(); //新建表格
            let row = sheet.get_collection_by_row(i.get_row_num());
            sheet_new.insert_new_row(&1, &1);
            let new_row = sheet_new.get_collection_by_row(&1);
            for i in row{
               let cell_value =  i.get_cell_value();
            }
        }
      
    }
    
}