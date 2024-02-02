use std::path::Path;
use umya_spreadsheet::*;
use std::fs;  
pub fn read_xlsx(filepath: &str) -> Spreadsheet {
    /*
        读取xlsx文件
        return: &SpredSheet
    */
    let file_path = Path::new(filepath); //文件目录
    reader::xlsx::read(file_path).expect("读取文件错误") //读取文件
}

pub fn book_set_header(sheet: &mut Worksheet, headers: &Vec<&Cell>, row_index: u32) {
    sheet.insert_new_row(&row_index, &row_index);
    for i in headers.clone() {
        sheet
            .get_cell_mut((*i.clone().get_coordinate().get_col_num(), row_index))
            .set_cell_value(i.get_cell_value().clone())
            .set_style(i.clone().get_style().clone());
    }
}

pub fn book_set_value(sheet: &mut Worksheet, rows: &Vec<&Cell>, row_index: u32) {
    sheet.insert_new_row(&row_index, &row_index);
    for i in rows.clone() {
        sheet
            .get_cell_mut((*i.clone().get_coordinate().get_col_num(), row_index))
            .set_cell_value(i.get_cell_value().clone())
            .set_style(i.clone().get_style().clone());
    }
}


pub fn get_cell_value(sheet: &Worksheet,col:u32,row:u32)-> String{
    sheet
    .get_cell((col, row))
    .unwrap()
    .get_cell_value()
    .get_raw_value()
    .to_string()
}

pub fn split_book(sheet: &Worksheet,header_number:u32,col_index:u32) {
    /*
      sheet:需要分离的表
      header_number: 表头行数
      col_index: 以对应列命名
    */
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
                book_set_header(sheet_new, &headers, s); //写入header
            }
            //插入第二行 写入数据
            let value_row_number = header_number + 1;
            book_set_value(sheet_new,&row,value_row_number);
//            sheet_new.insert_new_row(&value_row_number, &value_row_number);
//            for i in row.clone() {
//                sheet_new
//                    .get_cell_mut((*i.clone().get_coordinate().get_col_num(), value_row_number))
//                    .set_cell_value(i.get_cell_value().clone())
//                    .set_style(i.clone().get_style().clone());
//            }
            //获取对应文件名称
            let filename: String = get_cell_value(sheet,col_index,rowid);

            let filename = format!("./data/{name}.xlsx", name = filename);
            println!(">>>>>>>>>>>>>>>>>>>>>{filename}");
            let outpath = std::path::Path::new(&filename);
            writer::xlsx::write(&new_book, outpath).unwrap();
        }
    }
}



//检测data目录是否存在
pub fn check_path_exist(path:&str){
    let cpath = Path::new(path);
    if !cpath.exists(){
        fs::create_dir_all(path).expect(&format!("{path}"));
    }
}