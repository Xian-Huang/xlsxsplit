/*
    author:lidingyi
*/
use xlsxsplit::{split_main};
use xlsxsplit::select_file;
slint::include_modules!();


fn main() {
    //创建主页面
    let mainwidow = MainWindow::new().expect("MainWindow create fail!");
    mainwidow.on_selectfile(select_file);

    mainwidow.on_splitbook(split_main);
    //运行界面
    mainwidow.run().unwrap();
}