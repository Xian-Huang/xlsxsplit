/*
    author:lidingyi
*/
use xlsxsplit::{select_file,select_path};
use xlsxsplit::*;

fn main() {
    //创建主页面
    let mainwidow: MainWindow = MainWindow::new().expect("MainWindow create fail!");
    mainwidow.on_selectfile(select_file);
    mainwidow.on_selectpath(select_path);

    //装载分离功能
    presplitbook(&mainwidow);
    //运行界面
    mainwidow.run().unwrap();
}