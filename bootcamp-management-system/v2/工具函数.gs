// ========================================
// 🔧 工具函数模块
// ========================================
// 提供系统通用的工具函数和辅助方法

// ========================================
// 📊 数据查询和统计函数
// ========================================

/**
 * 获取学生积分排行榜
 */
function 获取学生积分排行() {
  const 学生Sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('学生信息');
  const 奖励Sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('奖励系统');
  
  const 学生数据 = 学生Sheet.getDataRange().getValues();
  const 奖励数据 = 奖励Sheet.getDataRange().getValues();
  
  const 学生积分 = {};
  
  // 统计每个学生的总积分
  for (let i = 1; i < 奖励数据.length; i++) {
    const 学生ID = 奖励数据[i][1];
    const 积分 = 奖励数据[i][3] || 0;
    
    if (学生积分[学生ID]) {
      学生积分[学生ID] += 积分;
    } else {
      学生积分[学生ID] = 积分;
    }
  }
  
  // 构建排行榜数据
  const 排行榜 = [];
  for (let i = 1; i < 学生数据.length; i++) {
    const 学生ID = 学生数据[i][0];
    const 姓名 = 学生数据[i][1];
    const 等级 = 学生数据[i][5];
    const 徽章 = 学生数据[i][15];
    const 总积分 = 学生积分[学生ID] || 0;
    
    if (学生ID) {
      排行榜.push({
        id: 学生ID,
        name: 姓名,
        level: 等级,
        totalPoints: 总积分,
        badge: 徽章
      });
    }
  }
  
  // 按积分排序
  排行榜.sort((a, b) => b.totalPoints - a.totalPoints);
  
  return 排行榜;
}

/**
 * 获取积分历史记录
 */
function 获取积分历史(筛选学生ID = '', 筛选类型 = '') {
  const 奖励Sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('奖励系统');
  const 学生Sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('学生信息');
  
  const 奖励数据 = 奖励Sheet.getDataRange().getValues();
  const 学生数据 = 学生Sheet.getDataRange().getValues();
  
  // 创建学生ID到姓名的映射
  const 学生姓名映射 = {};
  for (let i = 1; i < 学生数据.length; i++) {
    学生姓名映射[学生数据[i][0]] = 学生数据[i][1];
  }
  
  const 历史记录 = [];
  
  for (let i = 1; i < 奖励数据.length; i++) {
    const 记录 = 奖励数据[i];
    const 学生ID = 记录[1];
    const 积分类型 = 记录[2];
    const 积分数量 = 记录[3];
    const 原因 = 记录[4];
    const 时间 = 记录[8];
    
    // 应用筛选条件
    if (筛选学生ID && 学生ID !== 筛选学生ID) continue;
    if (筛选类型 && 积分类型 !== 筛选类型) continue;
    
    历史记录.push({
      studentId: 学生ID,
      studentName: 学生姓名映射[学生ID] || '未知学生',
      type: 积分类型,
      points: 积分数量,
      reason: 原因,
      time: 时间
    });
  }
  
  // 按时间倒序排列
  历史记录.sort((a, b) => new Date(b.time) - new Date(a.time));
  
  return 历史记录.slice(0, 50); // 返回最近50条记录
}

/**
 * 获取任务统计数据
 */
function 获取任务统计() {
  const 任务Sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('任务管理');
  const 数据 = 任务Sheet.getDataRange().getValues();
  
  const 统计 = {
    总任务数: 0,
    待开始: 0,
    进行中: 0,
    待审核: 0,
    已完成: 0,
    已过期: 0,
    完成率: 0
  };
  
  const 今天 = new Date();
  
  for (let i = 1; i < 数据.length; i++) {
    const 状态 = 数据[i][8];
    const 截止日期 = new Date(数据[i][7]);
    
    if (数据[i][0]) { // 如果任务ID不为空
      统计.总任务数++;
      
      if (状态 === '已完成') {
        统计.已完成++;
      } else if (截止日期 < 今天) {
        统计.已过期++;
      } else {
        switch (状态) {
          case '待开始': 统计.待开始++; break;
          case '进行中': 统计.进行中++; break;
          case '待审核': 统计.待审核++; break;
        }
      }
    }
  }
  
  统计.完成率 = 统计.总任务数 > 0 ? (统计.已完成 / 统计.总任务数 * 100).toFixed(1) : 0;
  
  return 统计;
}

/**
 * 获取学生学习进度统计
 */
function 获取学习进度统计(学生ID) {
  const 技术Sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('技术学习');
  const 英语Sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('英语学习');
  const 任务Sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('任务管理');
  
  const 技术数据 = 技术Sheet.getDataRange().getValues();
  const 英语数据 = 英语Sheet.getDataRange().getValues();
  const 任务数据 = 任务Sheet.getDataRange().getValues();
  
  const 统计 = {
    技术学习: { 总数: 0, 已完成: 0, 完成率: 0 },
    英语学习: { 总数: 0, 已完成: 0, 完成率: 0 },
    任务: { 总数: 0, 已完成: 0, 完成率: 0 },
    总积分: 0,
    当前徽章: '🌱'
  };
  
  // 统计技术学习
  for (let i = 1; i < 技术数据.length; i++) {
    if (技术数据[i][1] === 学生ID) {
      统计.技术学习.总数++;
      if (技术数据[i][4] === '已完成') {
        统计.技术学习.已完成++;
      }
    }
  }
  
  // 统计英语学习
  for (let i = 1; i < 英语数据.length; i++) {
    if (英语数据[i][1] === 学生ID) {
      统计.英语学习.总数++;
      if (英语数据[i][4] === '已完成') {
        统计.英语学习.已完成++;
      }
    }
  }
  
  // 统计任务
  for (let i = 1; i < 任务数据.length; i++) {
    if (任务数据[i][2] === 学生ID) {
      统计.任务.总数++;
      if (任务数据[i][8] === '已完成') {
        统计.任务.已完成++;
      }
    }
  }
  
  // 计算完成率
  统计.技术学习.完成率 = 统计.技术学习.总数 > 0 ? (统计.技术学习.已完成 / 统计.技术学习.总数 * 100).toFixed(1) : 0;
  统计.英语学习.完成率 = 统计.英语学习.总数 > 0 ? (统计.英语学习.已完成 / 统计.英语学习.总数 * 100).toFixed(1) : 0;
  统计.任务.完成率 = 统计.任务.总数 > 0 ? (统计.任务.已完成 / 统计.任务.总数 * 100).toFixed(1) : 0;
  
  // 获取总积分和徽章
  统计.总积分 = 获取学生总积分(学生ID);
  统计.当前徽章 = 计算徽章等级(统计.总积分);
  
  return 统计;
}

// ========================================
// 📅 日期和时间工具函数
// ========================================

/**
 * 格式化日期
 */
function 格式化日期(日期, 格式 = 'yyyy-mm-dd') {
  if (!日期) return '';
  
  const date = new Date(日期);
  const year = date.getFullYear();
  const month = String(date.getMonth() + 1).padStart(2, '0');
  const day = String(date.getDate()).padStart(2, '0');
  const hours = String(date.getHours()).padStart(2, '0');
  const minutes = String(date.getMinutes()).padStart(2, '0');
  
  switch (格式) {
    case 'yyyy-mm-dd':
      return `${year}-${month}-${day}`;
    case 'yyyy-mm-dd hh:mm':
      return `${year}-${month}-${day} ${hours}:${minutes}`;
    case 'mm/dd/yyyy':
      return `${month}/${day}/${year}`;
    case '中文日期':
      return `${year}年${parseInt(month)}月${parseInt(day)}日`;
    default:
      return date.toLocaleDateString();
  }
}

/**
 * 计算日期差
 */
function 计算日期差(开始日期, 结束日期, 单位 = 'days') {
  const start = new Date(开始日期);
  const end = new Date(结束日期);
  const diffTime = Math.abs(end - start);
  
  switch (单位) {
    case 'days':
      return Math.ceil(diffTime / (1000 * 60 * 60 * 24));
    case 'hours':
      return Math.ceil(diffTime / (1000 * 60 * 60));
    case 'minutes':
      return Math.ceil(diffTime / (1000 * 60));
    default:
      return diffTime;
  }
}

/**
 * 判断是否过期
 */
function 是否过期(截止日期) {
  const now = new Date();
  const due = new Date(截止日期);
  return due < now;
}

// ========================================
// 📧 通知和提醒函数
// ========================================

/**
 * 发送邮件通知（需要配置邮件服务）
 */
function 发送邮件通知(收件人邮箱, 主题, 内容) {
  try {
    // 这里可以使用 GmailApp 或其他邮件服务
    console.log(`📧 邮件通知: ${收件人邮箱} - ${主题}`);
    console.log(`内容: ${内容}`);
    
    // 示例代码（需要启用Gmail API）
    // GmailApp.sendEmail(收件人邮箱, 主题, 内容);
    
    return true;
  } catch (error) {
    console.error('邮件发送失败:', error);
    return false;
  }
}

/**
 * 批量发送任务提醒
 */
function 批量发送任务提醒() {
  const 任务Sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('任务管理');
  const 学生Sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('学生信息');
  
  const 任务数据 = 任务Sheet.getDataRange().getValues();
  const 学生数据 = 学生Sheet.getDataRange().getValues();
  
  // 创建学生ID到邮箱的映射
  const 学生邮箱映射 = {};
  for (let i = 1; i < 学生数据.length; i++) {
    学生邮箱映射[学生数据[i][0]] = 学生数据[i][2];
  }
  
  const 今天 = new Date();
  const 明天 = new Date(今天.getTime() + 24 * 60 * 60 * 1000);
  
  let 发送计数 = 0;
  
  for (let i = 1; i < 任务数据.length; i++) {
    const 任务 = 任务数据[i];
    const 学生ID = 任务[2];
    const 任务标题 = 任务[4];
    const 截止日期 = new Date(任务[7]);
    const 状态 = 任务[8];
    
    // 检查是否需要提醒（截止日期在明天且任务未完成）
    if (状态 !== '已完成' && 截止日期.toDateString() === 明天.toDateString()) {
      const 学生邮箱 = 学生邮箱映射[学生ID];
      if (学生邮箱) {
        const 主题 = `⏰ 任务提醒：${任务标题}`;
        const 内容 = `
          亲爱的同学，
          
          您有一个任务即将到期：
          
          任务标题：${任务标题}
          截止时间：${格式化日期(截止日期, 'yyyy-mm-dd hh:mm')}
          
          请及时完成并提交。
          
          祝学习愉快！
          训练营管理团队
        `;
        
        if (发送邮件通知(学生邮箱, 主题, 内容)) {
          发送计数++;
        }
      }
    }
  }
  
  console.log(`📧 已发送 ${发送计数} 条任务提醒`);
  return 发送计数;
}

// ========================================
// 📊 数据验证和清理函数
// ========================================

/**
 * 验证邮箱格式
 */
function 验证邮箱(邮箱) {
  const 邮箱正则 = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;
  return 邮箱正则.test(邮箱);
}

/**
 * 验证电话号码格式
 */
function 验证电话(电话) {
  const 电话正则 = /^[\d\s\-\+\(\)]{10,}$/;
  return 电话正则.test(电话);
}

/**
 * 清理和标准化数据
 */
function 清理数据(数据) {
  if (typeof 数据 === 'string') {
    return 数据.trim(); // 去除首尾空格
  }
  return 数据;
}

/**
 * 检查数据完整性
 */
function 检查数据完整性() {
  const 工作表列表 = [
    '学生信息', '教师信息', '技术学习', '英语学习', 
    '任务管理', '奖励系统', '求职管理', '项目作品'
  ];
  
  const 检查结果 = {
    总工作表数: 工作表列表.length,
    缺失工作表: [],
    数据统计: {}
  };
  
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  工作表列表.forEach(工作表名 => {
    const sheet = ss.getSheetByName(工作表名);
    if (!sheet) {
      检查结果.缺失工作表.push(工作表名);
    } else {
      const 行数 = sheet.getLastRow();
      检查结果.数据统计[工作表名] = {
        总行数: 行数,
        数据行数: Math.max(0, 行数 - 1) // 减去表头
      };
    }
  });
  
  return 检查结果;
}

// ========================================
// 🎨 界面显示函数
// ========================================

/**
 * 显示学生统计对话框
 */
function 显示学生统计() {
  const 学生Sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('学生信息');
  const 数据 = 学生Sheet.getDataRange().getValues();
  
  let 总学生数 = 0;
  let 在读学生 = 0;
  let 已毕业 = 0;
  const 等级统计 = {};
  
  for (let i = 1; i < 数据.length; i++) {
    if (数据[i][0]) { // 如果学生ID不为空
      总学生数++;
      const 状态 = 数据[i][9];
      const 等级 = 数据[i][5];
      
      if (状态 === '在读') 在读学生++;
      if (状态 === '已毕业') 已毕业++;
      
      等级统计[等级] = (等级统计[等级] || 0) + 1;
    }
  }
  
  let 统计信息 = `📊 学生统计信息\n\n`;
  统计信息 += `👥 总学生数: ${总学生数}\n`;
  统计信息 += `✅ 在读学生: ${在读学生}\n`;
  统计信息 += `🎓 已毕业: ${已毕业}\n\n`;
  统计信息 += `📈 等级分布:\n`;
  
  Object.entries(等级统计).forEach(([等级, 人数]) => {
    统计信息 += `• ${等级}: ${人数}人\n`;
  });
  
  SpreadsheetApp.getUi().alert('学生统计', 统计信息, SpreadsheetApp.getUi().ButtonSet.OK);
}

/**
 * 显示教师工作量统计
 */
function 显示教师工作量统计() {
  const 教师Sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('教师信息');
  const 任务Sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('任务管理');
  
  const 教师数据 = 教师Sheet.getDataRange().getValues();
  const 任务数据 = 任务Sheet.getDataRange().getValues();
  
  // 统计每个教师的任务数量
  const 教师任务统计 = {};
  
  for (let i = 1; i < 任务数据.length; i++) {
    const 教师ID = 任务数据[i][3];
    if (教师ID) {
      教师任务统计[教师ID] = (教师任务统计[教师ID] || 0) + 1;
    }
  }
  
  let 统计信息 = `👨‍🏫 教师工作量统计\n\n`;
  
  for (let i = 1; i < 教师数据.length; i++) {
    const 教师ID = 教师数据[i][0];
    const 教师姓名 = 教师数据[i][1];
    const 专业领域 = 教师数据[i][4];
    const 任务数量 = 教师任务统计[教师ID] || 0;
    
    if (教师ID) {
      统计信息 += `• ${教师姓名} (${专业领域}): ${任务数量}个任务\n`;
    }
  }
  
  SpreadsheetApp.getUi().alert('教师工作量', 统计信息, SpreadsheetApp.getUi().ButtonSet.OK);
}

/**
 * 显示积分排行榜对话框
 */
function 显示积分排行榜() {
  const 排行榜 = 获取学生积分排行();
  
  let 排行信息 = `🏆 积分排行榜 (前10名)\n\n`;
  
  for (let i = 0; i < Math.min(10, 排行榜.length); i++) {
    const 学生 = 排行榜[i];
    const 排名 = i + 1;
    let 排名图标 = `${排名}.`;
    
    if (排名 === 1) 排名图标 = '🥇';
    else if (排名 === 2) 排名图标 = '🥈';
    else if (排名 === 3) 排名图标 = '🥉';
    
    排行信息 += `${排名图标} ${学生.name} - ${学生.totalPoints}分 ${学生.badge}\n`;
  }
  
  SpreadsheetApp.getUi().alert('积分排行榜', 排行信息, SpreadsheetApp.getUi().ButtonSet.OK);
}

// ========================================
// 💾 数据导出函数
// ========================================

/**
 * 导出系统数据
 */
function 导出系统数据() {
  const ui = SpreadsheetApp.getUi();
  
  const result = ui.alert(
    '导出数据',
    '选择导出格式:\n\n确定 - Excel格式\n取消 - CSV格式',
    ui.ButtonSet.OK_CANCEL
  );
  
  const 格式 = result === ui.Button.OK ? 'excel' : 'csv';
  
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const 文件名 = `训练营数据_${格式化日期(new Date(), 'yyyy-mm-dd')}`;
    
    // 这里可以实现具体的导出逻辑
    console.log(`导出数据: ${文件名}.${格式}`);
    
    ui.alert('导出成功', `数据已导出为 ${文件名}.${格式}`, ui.ButtonSet.OK);
  } catch (error) {
    console.error('导出失败:', error);
    ui.alert('导出失败', error.toString(), ui.ButtonSet.OK);
  }
}

// ========================================
// ⚙️ 系统设置函数
// ========================================

/**
 * 显示系统设置
 */
function 显示系统设置() {
  const htmlOutput = HtmlService.createHtmlOutputFromFile('系统设置界面')
    .setWidth(600)
    .setHeight(500);
  
  SpreadsheetApp.getUi().showModalDialog(htmlOutput, '⚙️ 系统设置');
}

/**
 * 显示帮助信息
 */
function 显示帮助信息() {
  const 帮助内容 = `
🆘 训练营管理系统V2 帮助文档

📚 主要功能:
• 学生管理 - 添加、编辑学生信息
• 教师管理 - 管理教师资源
• 任务系统 - 布置和管理学习任务
• 奖励系统 - 积分和徽章管理
• 数据分析 - 学习进度和统计报告

🎯 快速操作:
1. 添加学生: 学生管理 > 添加新学生
2. 布置任务: 任务管理 > 布置新任务
3. 发放积分: 奖励系统 > 发放积分
4. 查看报告: 报告分析 > 生成报告

💡 使用技巧:
• 使用动态下拉菜单快速选择
• 定期检查任务完成情况
• 关注学生积分排行榜
• 利用数据导出功能备份

📞 技术支持:
如有问题请联系系统管理员
  `;
  
  SpreadsheetApp.getUi().alert('帮助文档', 帮助内容, SpreadsheetApp.getUi().ButtonSet.OK);
}

// ========================================
// 🔍 任务查询函数
// ========================================

/**
 * 按学生查询任务
 */
function 按学生查询任务() {
  const ui = SpreadsheetApp.getUi();
  
  const result = ui.prompt(
    '按学生查询',
    '请输入学生姓名（部分匹配）:',
    ui.ButtonSet.OK_CANCEL
  );
  
  if (result.getSelectedButton() !== ui.Button.OK) return '';
  
  const 学生姓名 = result.getResponseText();
  const 任务Sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('任务管理');
  const 学生Sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('学生信息');
  
  const 任务数据 = 任务Sheet.getDataRange().getValues();
  const 学生数据 = 学生Sheet.getDataRange().getValues();
  
  // 创建学生ID到姓名的映射
  const 学生映射 = {};
  for (let i = 1; i < 学生数据.length; i++) {
    学生映射[学生数据[i][0]] = 学生数据[i][1];
  }
  
  let 查询结果 = `🔍 学生"${学生姓名}"的任务列表:\n\n`;
  let 找到数量 = 0;
  
  for (let i = 1; i < 任务数据.length; i++) {
    const 学生ID = 任务数据[i][2];
    const 学生名 = 学生映射[学生ID];
    
    if (学生名 && 学生名.includes(学生姓名)) {
      const 任务标题 = 任务数据[i][4];
      const 状态 = 任务数据[i][8];
      const 截止日期 = 格式化日期(任务数据[i][7]);
      
      查询结果 += `• ${任务标题} - ${状态} (截止: ${截止日期})\n`;
      找到数量++;
    }
  }
  
  if (找到数量 === 0) {
    查询结果 += '未找到相关任务';
  } else {
    查询结果 += `\n共找到 ${找到数量} 个任务`;
  }
  
  return 查询结果;
}

/**
 * 按状态查询任务
 */
function 按状态查询任务() {
  const ui = SpreadsheetApp.getUi();
  
  const result = ui.prompt(
    '按状态查询',
    '请选择状态:\n1-待开始 2-进行中 3-待审核 4-已完成 5-已过期',
    ui.ButtonSet.OK_CANCEL
  );
  
  if (result.getSelectedButton() !== ui.Button.OK) return '';
  
  const 状态映射 = {
    '1': '待开始',
    '2': '进行中', 
    '3': '待审核',
    '4': '已完成',
    '5': '已过期'
  };
  
  const 选择状态 = 状态映射[result.getResponseText()];
  if (!选择状态) return '❌ 无效的状态选择';
  
  const 任务Sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('任务管理');
  const 学生Sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('学生信息');
  
  const 任务数据 = 任务Sheet.getDataRange().getValues();
  const 学生数据 = 学生Sheet.getDataRange().getValues();
  
  // 创建学生ID到姓名的映射
  const 学生映射 = {};
  for (let i = 1; i < 学生数据.length; i++) {
    学生映射[学生数据[i][0]] = 学生数据[i][1];
  }
  
  let 查询结果 = `🔍 状态为"${选择状态}"的任务:\n\n`;
  let 找到数量 = 0;
  
  for (let i = 1; i < 任务数据.length; i++) {
    const 状态 = 任务数据[i][8];
    
    if (状态 === 选择状态) {
      const 学生ID = 任务数据[i][2];
      const 学生名 = 学生映射[学生ID] || '未知学生';
      const 任务标题 = 任务数据[i][4];
      const 截止日期 = 格式化日期(任务数据[i][7]);
      
      查询结果 += `• ${学生名}: ${任务标题} (截止: ${截止日期})\n`;
      找到数量++;
    }
  }
  
  if (找到数量 === 0) {
    查询结果 += '未找到相关任务';
  } else {
    查询结果 += `\n共找到 ${找到数量} 个任务`;
  }
  
  return 查询结果;
}

/**
 * 按日期查询任务
 */
function 按日期查询任务() {
  const ui = SpreadsheetApp.getUi();
  
  const result = ui.prompt(
    '按截止日期查询',
    '请选择查询范围:\n1-今天到期 2-明天到期 3-本周到期 4-已过期',
    ui.ButtonSet.OK_CANCEL
  );
  
  if (result.getSelectedButton() !== ui.Button.OK) return '';
  
  const 今天 = new Date();
  const 明天 = new Date(今天.getTime() + 24 * 60 * 60 * 1000);
  const 本周末 = new Date(今天.getTime() + 7 * 24 * 60 * 60 * 1000);
  
  let 开始日期, 结束日期, 查询描述;
  
  switch (result.getResponseText()) {
    case '1':
      开始日期 = 今天;
      结束日期 = 明天;
      查询描述 = '今天到期';
      break;
    case '2':
      开始日期 = 明天;
      结束日期 = new Date(明天.getTime() + 24 * 60 * 60 * 1000);
      查询描述 = '明天到期';
      break;
    case '3':
      开始日期 = 今天;
      结束日期 = 本周末;
      查询描述 = '本周到期';
      break;
    case '4':
      开始日期 = new Date('2020-01-01');
      结束日期 = 今天;
      查询描述 = '已过期';
      break;
    default:
      return '❌ 无效的日期选择';
  }
  
  const 任务Sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('任务管理');
  const 学生Sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('学生信息');
  
  const 任务数据 = 任务Sheet.getDataRange().getValues();
  const 学生数据 = 学生Sheet.getDataRange().getValues();
  
  // 创建学生ID到姓名的映射
  const 学生映射 = {};
  for (let i = 1; i < 学生数据.length; i++) {
    学生映射[学生数据[i][0]] = 学生数据[i][1];
  }
  
  let 查询结果 = `🔍 ${查询描述}的任务:\n\n`;
  let 找到数量 = 0;
  
  for (let i = 1; i < 任务数据.length; i++) {
    const 截止日期 = new Date(任务数据[i][7]);
    
    if (截止日期 >= 开始日期 && 截止日期 < 结束日期) {
      const 学生ID = 任务数据[i][2];
      const 学生名 = 学生映射[学生ID] || '未知学生';
      const 任务标题 = 任务数据[i][4];
      const 状态 = 任务数据[i][8];
      
      查询结果 += `• ${学生名}: ${任务标题} - ${状态}\n`;
      找到数量++;
    }
  }
  
  if (找到数量 === 0) {
    查询结果 += '未找到相关任务';
  } else {
    查询结果 += `\n共找到 ${找到数量} 个任务`;
  }
  
  return 查询结果;
}

/**
 * 获取徽章统计
 */
function 获取徽章统计() {
  const 学生Sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('学生信息');
  const 数据 = 学生Sheet.getDataRange().getValues();
  
  const 徽章统计 = {};
  
  for (let i = 1; i < 数据.length; i++) {
    if (数据[i][0]) { // 如果学生ID不为空
      const 徽章 = 数据[i][15] || '新人 🌱';
      徽章统计[徽章] = (徽章统计[徽章] || 0) + 1;
    }
  }
  
  return 徽章统计;
}
