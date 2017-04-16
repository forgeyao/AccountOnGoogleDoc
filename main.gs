// 工作表“账目明细”里使用的背景色
var BACKGROUND_COLOR = ['#dd7e6b','#f9cb9c','#b4a7d6','#b6d7a8','#9fc5e8','#76a5af','#c27ba0','#f6b26b','#8e7cc3','#93c47d'];

/**
 * 羽毛球记账
 * 无论是场地费还是缴费，都执行本函数
 */
function tally() {
  // 获取“Input”工作表
  var sheet = SpreadsheetApp.getActive().getSheetByName('Input');
  // Logger.log('sheet id:' + sheet.getSheetId());

  var frozenrows = sheet.getFrozenRows();  // 前面冻结的行数
  var rownum = sheet.getLastRow();
  var colnum = sheet.getLastColumn();
  if(rownum <= frozenrows || colnum <= 1) {
    Logger.log('未输入或无法获取账目信息');
    return;
  }

  var range = sheet.getRange(frozenrows + 1, 1, rownum - frozenrows, colnum);
  var values = range.getValues();

  for(var i = 0; i < values.length; ++i) {
      if(values[i][1].trim() == '场地费' || values[i][1].trim() == '场租') {
          Logger.log('=======开始场地费结算=======');

          var users = values[i][2].split('、');
          var strnote = values[i][4] + ' 人均' + (values[i][3] / users.length).toFixed(2);

          siteFees(users, values[i][3]);
          writeDetailRecord(values[i][0], values[i][1], values[i][2], 0 - values[i][3], strnote);

          Logger.log('=======场地费结算结束=======');
      }
      else if(values[i][1].trim() == '缴费' || values[i][1].trim() == '充值') {
          Logger.log('=======开始缴费结算=======');

          var users = [values[i][2]];
          recharge(users, values[i][3]);
          writeDetailRecord(values[i][0], values[i][1], values[i][2], values[i][3], values[i][4]);

          Logger.log('=======缴费结算结束=======');
      }
      else{
          Logger.log('非法事件');
          return;
      }
  }

  // 清空输入数据
  Logger.log('清空输入数据：' + values.toString());
  range.clearContent();
}

/**
 * 场地费
 * 扣总费、扣参与人员平均费
 *
 * @param {Range} row 场地费记录，一行
 * @parm {double} fees 场地费
 */
function siteFees(users, fees) {
    // 扣总费
    var ret = totalExpenses('A3:C3', '会员卡资产', 0 - fees);
    if(ret < 0) {
        totalExpenses('A2:C2', '预存资产', ret);
    }

    // 扣个人
    var userfees = (0 - fees/users.length).toFixed(2);
    Logger.log("人均费：" + userfees.toString());
    userExpenses(users, userfees)
}

/**
 * 缴费（充值）
 *
 * @param {array} users 用户
 * @param {object} fees 费用
 */
function recharge(users, fees) {
    totalExpenses('A2:C2', '预存资产', fees);

    userExpenses(users, fees);
}


/**
 * 总体费用
 *
 * @param {String} location 位置
 * @param {string} title 名称
 * @param {double} fees 金额
 */
function totalExpenses(location, title, fees) {
    var sheet = SpreadsheetApp.getActive().getSheetByName('资产情况');
    var range = sheet.getRange(location);
    var log = title + "：" + range.getValues().toString() + " --> ";

    if(range.getCell(1, 1).getValue() != title) {
        Logger.log("error. not find " + title);
        return;
    }

    var cell = range.getCell(1, 2);
    var newvalue = cell.getValue() + fees;

    // 会员卡资产不能为负，
    // 会员卡资产不够时，先扣完，再从预存资产中扣
    if(newvalue < 0 && title == '会员卡资产') {
        if(cell.getValue() > 0) {
            cell.setValue(0);
        }
        Logger.log(log + range.getValues().toString());
        return newvalue;
    }

    cell.setValue(newvalue);
    cell.setFontColor(newvalue < 0 ? 'red' : 'black');

    Logger.log(log + range.getValues().toString());
    return 0;
}

/**
 * 个人账户
 *
 * @param {} users 用户
 * @param {} fees 费用金额
 */
function userExpenses(users, fees) {
    var sheet = SpreadsheetApp.getActive().getSheetByName('资产情况');
    var frozenrows = sheet.getFrozenRows();
    var rownum = sheet.getLastRow();
    var colnum = sheet.getLastColumn();

    if(rownum == frozenrows) {
        for(var i in users) {
            addUserAccount(sheet, users[i], fees);
        }
        return;
    }

    var range = sheet.getRange(frozenrows + 1, 1, rownum - frozenrows, colnum);
    var userdata = range.getValues();

    var userarray = [];
    for (var i = 0; i < userdata.length; ++i) {
        userarray.push(userdata[i][0].trim());
    }
//    Logger.log(userarray.toString());

    for(var i in users) {
        var index = userarray.indexOf(users[i].trim());
        if(index != -1) {    // find
              var cell = range.getCell(index + 1, 2);
              var olddata = Number(cell.getValue());
              var newdata = olddata + Number(fees);

              cell.setValue(newdata)
              cell.setFontColor(newdata < 0 ? 'red' : 'black');

              Logger.log("用户:" + users[i] + " 金额更新 " + olddata + " -> " + newdata);
          }
          else {   // not find
              addUserAccount(sheet, users[i], fees);
          }
    }
}

/**
 * 新增用户
 */
function addUserAccount(sheet, user, fees){
  sheet.appendRow([user, fees]);
  var cell = sheet.getRange(sheet.getLastRow(), 2).getCell(1, 1);
  cell.setFontColor(fees < 0 ? 'red' : 'black')
      .setFontWeight('bold');
  Logger.log("新增用户 :" + user + ',' + fees);
}

/**
 * 写消费记录
 *
 * @param {} time  时间
 * @param {} event 事件(场地费／缴费)
 * @param {} name  涉及用户姓名,多用户时已处理
 * @param {} fees  费用，消费为负值，充值为正
 * @param {} note  备注
 */
function writeDetailRecord(time, event, name, fees, note) {
    var sheet = SpreadsheetApp.getActive().getSheetByName('账目明细');
    var frozenrows = sheet.getFrozenRows();

    var id = 1;
    var date = '';
    var color = '';
    if(sheet.getLastRow() > frozenrows) {
        var beforerange = sheet.getRange(frozenrows + 1, 1, 1, 2);
        id = beforerange.getCell(1, 1).getValue() + 1;
        date = beforerange.getCell(1, 2).getValue().toDateString();
        color = beforerange.getBackground();
    }

    if(date != time.toDateString()) {
        var newcolor = BACKGROUND_COLOR[id % 10];
        if(newcolor == color) {
            color = BACKGROUND_COLOR[(id + 1) % 10];
        }
        else {
            color = newcolor;
        }
    }

    var row = [id, time, event, name, fees, note];
    var array = [];
    array.push(row);

    sheet.insertRowAfter(frozenrows);

    var range = sheet.getRange(frozenrows + 1, 1, 1, 6);
    range.setValues(array);
    range.setBackground(color);
    Logger.log("账目：" + row.toString());
}

