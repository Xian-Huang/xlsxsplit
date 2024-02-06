/*
    author:lidingyi
*/
use xlsxsplit::*;
use xlsxsplit::{select_file, select_path};

fn main() {
    //创建主页面
    let mainwidow: MainWindow = MainWindow::new().expect("MainWindow create fail!");
    let weak = mainwidow.as_weak().unwrap();
    mainwidow.on_selectfile(move || select_file(&weak));
    let weak = mainwidow.as_weak().unwrap();
    mainwidow.on_selectpath(move || select_path(&weak));

    //装载分离功能
    presplitbook(&mainwidow);
    //运行界面
    mainwidow.run().unwrap();
}
