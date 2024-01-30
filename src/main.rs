/*
author:lidingyi
*/
use umya_spreadsheet::*;
use core::panic;
use std::path::Path;




fn main(){
    let file_path = Path::new("./test.xlsx"); //文件目录
    let header_number = 1u32;
    let col_index = 0;
    let book = reader::xlsx::read(file_path).unwrap(); //读取文件
    let sheet = book.get_sheet_collection().first().unwrap(); //获取第一个sheet
    let headers = sheet.get_collection_by_row(&header_number);
    
    //循环每一行
    for i in sheet.get_row_dimensions(){
        let rowid = *i.get_row_num();
        if rowid>header_number{
            let mut new_book = new_file(); //创建新的文件
            let sheet_new = new_book.get_sheet_by_name_mut("Sheet1").unwrap(); //新建表格
            let row = sheet.get_collection_by_row(i.get_row_num());
            //插入第一行 将header写入
            sheet_new.insert_new_row(&1, &1);
            for i in headers.clone(){
                sheet_new.get_cell_mut((*i.clone().get_coordinate().get_col_num(),1)).set_cell_value(i.get_cell_value().clone());
            }
            //插入第二行 写入数据
            sheet_new.insert_new_row(&2, &2);
            for i in row.clone(){
                sheet_new.get_cell_mut((*i.clone().get_coordinate().get_col_num(),2)).set_cell_value(i.get_cell_value().clone());
            }
            //写入文件
            let filename = sheet.get_cell((1,rowid)).unwrap().get_cell_value().get_raw_value().to_string();

            let filename = format!("./data/{name}.xlsx",name=filename);
            println!(">>>>>>>>>>>>>>>>>>>>>{filename}");
            let outpath = std::path::Path::new(&filename);
            writer::xlsx::write(&new_book, outpath).unwrap();

        }
      
    }
    
}