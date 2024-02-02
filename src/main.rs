/*
author:lidingyi
*/
use xlsxsplit::{read_xlsx,split_book,check_path_exist};

fn main() {
    let header_number = 2u32; //header 行数 前几行
    let col_index = 1;
    let book = &read_xlsx("./test.xlsx"); //读取文件
    let sheet = book.get_sheet_collection().first().unwrap(); //获取第一个sheet
    
    check_path_exist("./data");
    split_book(sheet, header_number,col_index);
}
