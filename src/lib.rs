use rfd::FileDialog;
use slint::SharedString;
use std::fs;
use std::path::Path;
use umya_spreadsheet::*;
slint::include_modules!();


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

pub fn get_cell_value(sheet: &Worksheet, col: u32, row: u32) -> String {
    sheet
        .get_cell((col, row))
        .unwrap()
        .get_cell_value()
        .get_raw_value()
        .to_string()
}

pub fn split_book(sheet: &Worksheet, header_number: i32, col_index: i32, output: &Path,mainwindow:&MainWindow) {
    /*
      sheet:需要分离的表
      header_number: 表头行数
      col_index: 以对应列命名
    */
    //循环每一行
    let mut finished = 0;
    let sum_row = sheet.get_highest_row();
    for i in sheet.get_row_dimensions() {
        let rowid = *i.get_row_num();
        if rowid > header_number.try_into().unwrap() {
            let mut new_book = new_file(); //创建新的文件
            let sheet_new = new_book.get_sheet_by_name_mut("Sheet1").unwrap(); //新建表格
            let row = sheet.get_collection_by_row(i.get_row_num());
            //插入{header_number}行 将header写入
            for s in 1..=header_number {
                let headers = sheet.get_collection_by_row(&s.try_into().unwrap());
                book_set_header(sheet_new, &headers, s.try_into().unwrap()); //写入header
            }
            //插入第二行 写入数据
            let value_row_number = header_number + 1;
            book_set_value(sheet_new, &row, value_row_number.try_into().unwrap());
            //            sheet_new.insert_new_row(&value_row_number, &value_row_number);
            //            for i in row.clone() {
            //                sheet_new
            //                    .get_cell_mut((*i.clone().get_coordinate().get_col_num(), value_row_number))
            //                    .set_cell_value(i.get_cell_value().clone())
            //                    .set_style(i.clone().get_style().clone());
            //            }
            //获取对应文件名称
            let filename: String = get_cell_value(sheet, col_index.try_into().unwrap(), rowid);
            let filepath = format!("{filename}.xlsx");
            let ouput_file = output.join(filepath.clone());
            let outpath = std::path::Path::new(&ouput_file);
            if outpath.exists(){
                if let Err(res) = writer::xlsx::write(&new_book, outpath){
                    mainwindow.invoke_error_window_show("写入文件失败！".into());
                    break;
                }
            }else{
                mainwindow.invoke_error_window_show("输出目录不存在".into());
            }
            
            finished +=1;
            let pregress = finished as f32/sum_row as f32;
            println!(">>>>>>>>>>>>>>>>>>>>>{:?} {pregress} {finished}/{sum_row}", ouput_file.as_path());
            mainwindow.set_progress(pregress);
            
        }
    }
}

//检测data目录是否存在
pub fn check_path_exist(path: &str) -> &Path {
    let cpath = Path::new(path);
    if !cpath.exists() {
        fs::create_dir_all(path).expect(&format!("{path}"));
    }
    cpath
}

pub fn split_main(input:SharedString,output:SharedString,header_number:i32,col_index:i32,mainwindow:&MainWindow) {
    
    let book = &read_xlsx(input.as_str()); //读取文件
    let sheet = book.get_sheet_collection().first().unwrap(); //获取第一个sheet
    let path = check_path_exist(output.as_str());
    split_book(sheet, header_number, col_index, path,mainwindow);
    mainwindow.invoke_disable_btn();
}


pub fn presplitbook(mainwindow:&MainWindow){
    let weak = mainwindow.as_weak().unwrap();
    mainwindow.on_splitbook(move || {
        split_main(weak.get_input(), weak.get_output(), weak.get_header_number(), weak.get_col_index(), &weak)
    });
}

//选择文件
pub fn select_file(mainwindow:&MainWindow)->SharedString {
    let files = FileDialog::new()
        .add_filter("data", &["csv", "xlsx"])
        .set_directory("/")
        .pick_file();
    if let Some(files) = files{
            println!("{:?}",files.to_str().unwrap());
            return files.to_str().unwrap().to_string().into();
    }else {
        println!("选择输入文件失败");
        mainwindow.invoke_error_window_show("选择输入文件失败".into());
        return SharedString::new();
    }
}


//选择输出目录
pub fn select_path(mainwindow:&MainWindow)->SharedString {
    let files = FileDialog::new()
        .pick_folder();
    if let Some(files) = files{
            println!("{:?}",files.to_str().unwrap());
            return files.to_str().unwrap().to_string().into();
    }else {
        println!("选择输出目录失败");
        mainwindow.invoke_error_window_show("选择输出目录失败".into());
        return SharedString::new();
    }
}
