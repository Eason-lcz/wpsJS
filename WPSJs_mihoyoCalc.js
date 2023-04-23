var btnclickCount = 2;
/**
 * btn点击触发
 */
function CommandButton2_Click() {
    var cells1 = Application.ActiveWorkbook.Sheets.Item("Sheet1").Cells;
    if (btnclickCount > 2) {
        cells1.Item(btnclickCount, 1).Formula = '';
        cells1.Item(btnclickCount, 2).Formula = '';
        cells1.Item(btnclickCount, 3).Formula = '';
        cells1.Item(btnclickCount, 4).Formula = '';
        btnclickCount -= 1;
        alert("回滚成功至" + btnclickCount + "行");
    } else {
        btnclickCount = 2
        alert("无需清除");
    }

}
function CommandButton1_Click() {
    var x = btnclickCount;
    write(x);
    btnclickCount += 1;
}
function write(x) {
    // 获取单元格值
    alert("查看第" + x + "行数据");
    var cells = Application.ActiveWorkbook.Sheets.Item("Sheet1").Cells;
    var jin = Number(cells.Item(x, 1).Formula);
    var zi = Number(cells.Item(x, 2).Formula);
    var lan = Number(cells.Item(x, 3).Formula);
    var lv = Number(cells.Item(x, 4).Formula);
    var oldArry = [];
    oldArry = [jin, zi, lan, lv];
    alert(oldArry.toString());
    alert("转换前的 \n jin:" + jin + ",zi:" + zi + ",lan:" + lan + ",lv:" + lv)
    if (jin != '' && zi != '' && lan != '' && lv != '') {
        if (lv >= 3) {
            var num1 = Math.floor(lv / 3);
            lv = lv % 3;
            lan += Number(num1);
        }
        if (lan >= 3) {
            var num2 = Math.floor(lan / 3);
            lan = lan % 3;
            zi += Number(num2);
        }
        if (zi >= 3) {
            var num3 = Math.floor(zi / 3);
            zi = zi % 3;
            jin += Number(num3);
        }
        alert("转换后的 \n  jin:" + jin + ",zi:" + zi + ",lan:" + lan + ",lv:" + lv)
        var newArry = [];
        newArry = [jin, zi, lan, lv];
        alert(newArry.toString());
        //   写入数据
        if (oldArry != newArry) {
            var RangNum = x + 1;
            cells.Item(RangNum, 1).Formula = jin;
            cells.Item(RangNum, 2).Formula = zi;
            cells.Item(RangNum, 3).Formula = lan;
            cells.Item(RangNum, 4).Formula = lv;
        }else{
            alert("数据无变化，无需转换")
        }
    } else {
        alert("暂无数据")
    }
}
