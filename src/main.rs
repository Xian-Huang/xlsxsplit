/*
author:lidingyi
*/
use xlsxsplit::{read_xlsx,split_book,check_path_exist};

fn main() {
    let header_number = 2u32; //header 行数 前几行
    let col_index = 1;
    let output = "./data";
    let book = &read_xlsx("./test.xlsx"); //读取文件
    let sheet = book.get_sheet_collection().first().unwrap(); //获取第一个sheet

    let path = check_path_exist(output);
    split_book(sheet, header_number,col_index,path);
}
