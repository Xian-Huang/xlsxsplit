/*
    author:lidingyi
*/
use xlsxsplit::select_file;
use xlsxsplit::*;
fn main() {
    //创建主页面
    let mainwidow: MainWindow = MainWindow::new().expect("MainWindow create fail!");
    mainwidow.on_selectfile(select_file);
    //装载分离功能
    presplitbook(&mainwidow);
    //运行界面
    mainwidow.run().unwrap();
}