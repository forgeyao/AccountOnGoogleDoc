
// #ffe599,#a4c2f4,#ea9999,#cfe2f3,#e6b8af,#fce5cd,#d5a6bd,#783f04,#85200c,#4a86e8
var BACKGROUND_COLOR = ['#dd7e6b','#f9cb9c','#b4a7d6','#b6d7a8','#9fc5e8','#76a5af','#c27ba0','#f6b26b','#8e7cc3','#93c47d'];

/**
 * 羽毛球记账
 *
 */
function tallyBadminton() {
  var sheet = SpreadsheetApp.getActive().getSheetByName('Input');
//  Logger.log('sheet id:' + sheet.getSheetId());

  var frozenrows = sheet.getFrozenRows();  // 前面冻结的行数
  var range = sheet.getRange(frozenrows + 1, 1, sheet.getLastRow() - frozenrows, sheet.getLastColumn());
  var values = range.getValues();

  for(var i = 0; i < values.length; ++i) {
      if(values[i][1].trim() == '场地费' || values[i][1].trim() == '场租') {
          Logger.log('=======开始场地费结算=======');

          var users = values[i][2].split('、');
          var struser = String();
          for(var j in users) {
              struser += users[j] + '\n';
          }

          var strnote = values[i][4] + '\n人均' + values[i][3] / users.length;

          siteFees(users, values[i][3]);

          writeDetailRecord(values[i][0], values[i][1], struser, 0 - values[i][3], strnote);

          Logger.log("=======场地费结算结束=======");
      }
      else if(values[i][1].trim() == '缴费' || values[i][1].trim() == '充值') {
          Logger.log('=======开始缴费结算=======');

          var users = [values[i][2]];
          recharge(users, values[i][3]);

          writeDetailRecord(values[i][0], values[i][1], values[i][2], values[i][3], values[i][4]);

          Logger.log('=======缴费结算结束=======');
      }
  }

  // 清空输入数据
  Logger.log('清空输入数据：' + values.toString());
  range.clearContent();
}

/**
 * 场地费
 * 扣总费、扣参与人员平均费、写记录
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
    userExpenses(users, 0 - fees / users.length)
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
    Logger.log(title + "前：" + range.getValues().toString());

    if(range.getCell(1, 1).getValue() == title) {
        var cell = range.getCell(1, 2);
        var newvalue = cell.getValue() + fees;

        // 会员卡资产不能为负，
        // 会员卡资产不够时，先扣完，再从预存资产中扣
        if(newvalue < 0 && title == '会员卡资产') {
            if(cell.getValue() > 0) {
                range.getCell(1, 3).setValue(cell.getValue());
                cell.setValue(0);
            }
            Logger.log(title + "后：" + range.getValues().toString());
            return newvalue;
        }

        range.getCell(1, 3).setValue(cell.getValue())
                           .setFontColor('black');

        var newvalue = cell.getValue() + fees;
        cell.setValue(newvalue);
        if(newvalue < 0) {
            cell.setFontColor('red');
        }
        else {
            cell.setFontColor('black');
        }

        Logger.log(title + "后：" + range.getValues().toString());
    }
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
    var range = sheet.getRange(frozenrows + 1, 1, sheet.getLastRow() - frozenrows, sheet.getLastColumn());
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
              var olddata = cell.getValue();
              var newdata = olddata + fees;

              range.getCell(index + 1, 3).setValue(olddata)
                                         .setFontColor('black');
              cell.setValue(newdata)
              if(newdata < 0) {
                  cell.setFontColor('red');
              }
              else {
                  cell.setFontColor('black');
              }

              Logger.log("用户:" + users[i] + " 金额更新 " + olddata + " -> " + newdata);
          }
          else {   // not find
              sheet.appendRow([users[i], fees, 0]);
              var cell = sheet.getRange(sheet.getLastRow(), 2).getCell(1, 1);

              if(fees < 0) {
                  cell.setFontColor('red')
                      .setFontWeight('bold');
              }
              else {
                  cell.setFontColor('black')
                      .setFontWeight('bold');
              }
              Logger.log("新增用户 :" + users[i] + ',' + fees + ',' + 0);
          }
    }
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
    Logger.log("write record:" + row.toString());
}

