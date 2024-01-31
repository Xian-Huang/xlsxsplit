/*
author:lidingyi
*/
use std::path::Path;
use umya_spreadsheet::*;

fn main() {
    let file_path = Path::new("./test.xlsx"); //文件目录
    let header_number = 2u32; //header 行数 前几行
    let col_index = 1;
    let book = reader::xlsx::read(file_path).unwrap(); //读取文件
    let sheet = book.get_sheet_collection().first().unwrap(); //获取第一个sheet

    //循环每一行
    
    for i in sheet.get_row_dimensions() {
        let rowid = *i.get_row_num();
        if rowid > header_number {
            let mut new_book = new_file(); //创建新的文件
            let sheet_new = new_book.get_sheet_by_name_mut("Sheet1").unwrap(); //新建表格
            let row = sheet.get_collection_by_row(i.get_row_num());
            //插入{header_number}行 将header写入
            for s in 1..=header_number {
                let headers = sheet.get_collection_by_row(&s);
                sheet_new.insert_new_row(&s, &s);
                for i in headers.clone() {
                    sheet_new
                        .get_cell_mut((*i.clone().get_coordinate().get_col_num(), s))
                        .set_cell_value(i.get_cell_value().clone()).set_style(i.clone().get_style().clone());
                }
            }
            //插入第二行 写入数据
            let value_row_number = header_number + 1;
            sheet_new.insert_new_row(&value_row_number, &value_row_number);
            for i in row.clone() {
                sheet_new
                    .get_cell_mut((*i.clone().get_coordinate().get_col_num(), value_row_number))
                    .set_cell_value(i.get_cell_value().clone()).set_style(i.clone().get_style().clone());
            }
            //写入文件
            let filename = sheet
                .get_cell((col_index, rowid))
                .unwrap()
                .get_cell_value()
                .get_raw_value()
                .to_string();

            let filename = format!("./data/{name}.xlsx", name = filename);
            println!(">>>>>>>>>>>>>>>>>>>>>{filename}");
            let outpath = std::path::Path::new(&filename);
            writer::xlsx::write(&new_book, outpath).unwrap();
        }
    }
}
