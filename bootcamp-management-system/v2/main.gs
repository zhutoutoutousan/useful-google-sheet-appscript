// ========================================
// 🎓 IT前端训练营管理系统 V2.0
// ========================================
// 全中文界面，现代化设计，智能任务管理
// 特色功能：动态下拉菜单、任务系统、奖励机制、HTML界面

// ========================================
// 📊 系统初始化设置
// ========================================

/**
 * 初始化整个训练营管理系统V2
 */
function 初始化训练营系统V2() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  try {
    // 创建所有必要的工作表
    创建工作表如果不存在(ss, '仪表板');
    创建工作表如果不存在(ss, '学生信息');
    创建工作表如果不存在(ss, '教师信息');
    创建工作表如果不存在(ss, '技术学习');
    创建工作表如果不存在(ss, '英语学习');
    创建工作表如果不存在(ss, '任务管理');
    创建工作表如果不存在(ss, '奖励系统');
    创建工作表如果不存在(ss, '求职管理');
    创建工作表如果不存在(ss, '项目作品');
    创建工作表如果不存在(ss, '考核评估');
    创建工作表如果不存在(ss, '报告分析');
    创建工作表如果不存在(ss, '系统设置');
    
    // 初始化所有工作表结构
    设置仪表板();
    设置学生信息表();
    设置教师信息表();
    设置技术学习表();
    设置英语学习表();
    设置任务管理表();
    设置奖励系统表();
    设置求职管理表();
    设置项目作品表();
    设置考核评估表();
    设置报告分析表();
    设置系统设置表();
    
    // 创建现代化仪表板
    创建现代化仪表板();
    
    // 添加示例数据
    添加示例数据();
    
    SpreadsheetApp.getUi().alert(
      '🎉 系统初始化完成！', 
      '训练营管理系统V2已成功初始化！\n\n✨ 新功能已就绪：\n• 中文界面\n• 动态下拉菜单\n• 智能任务系统\n• 奖励管理\n• 现代化界面', 
      SpreadsheetApp.getUi().ButtonSet.OK
    );
  } catch (error) {
    console.error('初始化过程中发生错误:', error);
    SpreadsheetApp.getUi().alert('❌ 初始化失败', `错误信息: ${error.toString()}`, SpreadsheetApp.getUi().ButtonSet.OK);
  }
}

/**
 * 创建工作表（如果不存在）
 */
function 创建工作表如果不存在(ss, 工作表名称) {
  if (!ss.getSheetByName(工作表名称)) {
    const sheet = ss.insertSheet(工作表名称);
    console.log(`已创建工作表: ${工作表名称}`);
    return sheet;
  }
  return ss.getSheetByName(工作表名称);
}

// ========================================
// 👥 学生信息管理
// ========================================

/**
 * 设置学生信息表结构
 */
function 设置学生信息表() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('学生信息');
  
  // 清除现有数据
  sheet.clear();
  
  // 设置表头
  const headers = [
    '学生ID', '姓名', '邮箱', '电话', '入学日期', '当前等级',
    '技术分数', '英语分数', '求职准备度', '状态', '指导老师',
    'GitHub', 'LinkedIn', '作品集', '总积分', '当前徽章', '备注', '创建时间'
  ];
  
  const headerRange = sheet.getRange(1, 1, 1, headers.length);
  headerRange.setValues([headers]);
  
  // 美化表头
  headerRange.setFontWeight('bold');
  headerRange.setBackground('#1a73e8');
  headerRange.setFontColor('white');
  headerRange.setHorizontalAlignment('center');
  headerRange.setFontSize(11);
  
  // 设置数据验证
  const 状态选项 = ['在读', '已毕业', '暂停学习', '退学'];
  const 等级选项 = ['初级', '中级', '高级', '专家'];
  
  // 状态下拉列表
  const statusRange = sheet.getRange(2, 10, 1000, 1);
  statusRange.setDataValidation(SpreadsheetApp.newDataValidation()
    .requireValueInList(状态选项, true)
    .build());
  
  // 等级下拉列表
  const levelRange = sheet.getRange(2, 6, 1000, 1);
  levelRange.setDataValidation(SpreadsheetApp.newDataValidation()
    .requireValueInList(等级选项, true)
    .build());
  
  // 设置格式
  sheet.getRange('E:E').setNumberFormat('yyyy-mm-dd');
  sheet.getRange('G:I').setNumberFormat('0.0');
  sheet.getRange('O:O').setNumberFormat('#,##0');
  sheet.getRange('R:R').setNumberFormat('yyyy-mm-dd hh:mm:ss');
  
  // 冻结表头
  sheet.setFrozenRows(1);
  
  // 自动调整列宽
  sheet.autoResizeColumns(1, headers.length);
  
  // 添加条件格式
  添加学生分数条件格式(sheet);
}

/**
 * 添加学生分数的条件格式
 */
function 添加学生分数条件格式(sheet) {
  const scoreRange = sheet.getRange('G2:I1000');
  
  const rules = [
    SpreadsheetApp.newConditionalFormatRule()
      .whenNumberLessThan(60)
      .setBackground('#fce8e6')
      .setFontColor('#d93025')
      .setRanges([scoreRange])
      .build(),
    SpreadsheetApp.newConditionalFormatRule()
      .whenNumberBetween(60, 79)
      .setBackground('#fef7e0')
      .setFontColor('#f9ab00')
      .setRanges([scoreRange])
      .build(),
    SpreadsheetApp.newConditionalFormatRule()
      .whenNumberGreaterThanOrEqualTo(80)
      .setBackground('#e8f5e8')
      .setFontColor('#137333')
      .setRanges([scoreRange])
      .build()
  ];
  
  sheet.setConditionalFormatRules(rules);
}

/**
 * 添加新学生
 */
function 添加新学生(姓名, 邮箱, 电话, 入学日期, 当前等级, 指导老师) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('学生信息');
  const 今天 = new Date();
  const 学生ID = '学生' + Utilities.getUuid().substring(0, 8).toUpperCase();
  
  // 将入学日期字符串转换为Date对象
  const 入学日期对象 = typeof 入学日期 === 'string' ? new Date(入学日期) : 入学日期;
  
  const 新行数据 = [
    学生ID,           // 学生ID
    姓名,             // 姓名
    邮箱,             // 邮箱
    电话,             // 电话
    入学日期对象,      // 入学日期
    当前等级,          // 当前等级
    0,               // 技术分数
    0,               // 英语分数
    0,               // 求职准备度
    '在读',           // 状态
    指导老师,          // 指导老师
    '',              // GitHub
    '',              // LinkedIn
    '',              // 作品集
    0,               // 总积分
    '新人 🌱',        // 当前徽章
    '',              // 备注
    今天              // 创建时间
  ];
  
  sheet.appendRow(新行数据);
  
  // 更新仪表板
  更新仪表板();
  
  return 学生ID;
}

// ========================================
// 👨‍🏫 教师信息管理
// ========================================

/**
 * 设置教师信息表结构
 */
function 设置教师信息表() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('教师信息');
  
  // 清除现有数据
  sheet.clear();
  
  // 设置表头
  const headers = [
    '教师ID', '姓名', '邮箱', '电话', '专业领域', '工作经验年限',
    '负责学生数', '状态', '时薪', '可用性', '教学评分', '备注'
  ];
  
  const headerRange = sheet.getRange(1, 1, 1, headers.length);
  headerRange.setValues([headers]);
  
  // 美化表头
  headerRange.setFontWeight('bold');
  headerRange.setBackground('#34a853');
  headerRange.setFontColor('white');
  headerRange.setHorizontalAlignment('center');
  headerRange.setFontSize(11);
  
  // 设置数据验证
  const 专业领域选项 = ['前端开发', '后端开发', '全栈开发', '数据库', 'DevOps', '英语教学', '职业规划'];
  const 状态选项 = ['在职', '休假', '离职'];
  const 可用性选项 = ['可用', '忙碌', '不可用'];
  
  // 专业领域下拉列表
  const specRange = sheet.getRange(2, 5, 1000, 1);
  specRange.setDataValidation(SpreadsheetApp.newDataValidation()
    .requireValueInList(专业领域选项, true)
    .build());
  
  // 状态下拉列表
  const statusRange = sheet.getRange(2, 8, 1000, 1);
  statusRange.setDataValidation(SpreadsheetApp.newDataValidation()
    .requireValueInList(状态选项, true)
    .build());
  
  // 可用性下拉列表
  const availRange = sheet.getRange(2, 10, 1000, 1);
  availRange.setDataValidation(SpreadsheetApp.newDataValidation()
    .requireValueInList(可用性选项, true)
    .build());
  
  // 设置格式
  sheet.getRange('F:F').setNumberFormat('0');
  sheet.getRange('G:G').setNumberFormat('0');
  sheet.getRange('I:I').setNumberFormat('¥#,##0');
  sheet.getRange('K:K').setNumberFormat('0.0');
  
  // 冻结表头
  sheet.setFrozenRows(1);
  
  // 自动调整列宽
  sheet.autoResizeColumns(1, headers.length);
}

/**
 * 添加新教师
 */
function 添加新教师(姓名, 邮箱, 电话, 专业领域, 工作经验年限, 时薪) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('教师信息');
  const 教师ID = '教师' + Utilities.getUuid().substring(0, 8).toUpperCase();
  
  const 新行数据 = [
    教师ID,            // 教师ID
    姓名,              // 姓名
    邮箱,              // 邮箱
    电话,              // 电话
    专业领域,           // 专业领域
    工作经验年限,        // 工作经验年限
    0,                // 负责学生数
    '在职',            // 状态
    时薪,              // 时薪
    '可用',            // 可用性
    5.0,              // 教学评分
    ''                // 备注
  ];
  
  sheet.appendRow(新行数据);
  
  return 教师ID;
}

// ========================================
// 📚 技术学习管理
// ========================================

/**
 * 设置技术学习表结构
 */
function 设置技术学习表() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('技术学习');
  
  // 清除现有数据
  sheet.clear();
  
  // 设置表头
  const headers = [
    '学习ID', '学生ID', '课程名称', '技术分类', '状态', '进度百分比', '分数',
    '开始日期', '完成日期', '指导老师', '学习资源', '作业链接', '备注'
  ];
  
  const headerRange = sheet.getRange(1, 1, 1, headers.length);
  headerRange.setValues([headers]);
  
  // 美化表头
  headerRange.setFontWeight('bold');
  headerRange.setBackground('#ff6d01');
  headerRange.setFontColor('white');
  headerRange.setHorizontalAlignment('center');
  headerRange.setFontSize(11);
  
  // 设置数据验证
  const 技术分类选项 = ['HTML/CSS', 'JavaScript', 'React', 'Vue.js', 'Node.js', '数据库', 'DevOps', '测试', '项目实战'];
  const 状态选项 = ['未开始', '学习中', '已完成', '需要复习'];
  
  // 技术分类下拉列表
  const categoryRange = sheet.getRange(2, 4, 1000, 1);
  categoryRange.setDataValidation(SpreadsheetApp.newDataValidation()
    .requireValueInList(技术分类选项, true)
    .build());
  
  // 状态下拉列表
  const statusRange = sheet.getRange(2, 5, 1000, 1);
  statusRange.setDataValidation(SpreadsheetApp.newDataValidation()
    .requireValueInList(状态选项, true)
    .build());
  
  // 设置格式
  sheet.getRange('F:F').setNumberFormat('0%');
  sheet.getRange('G:G').setNumberFormat('0.0');
  sheet.getRange('H:I').setNumberFormat('yyyy-mm-dd');
  
  // 冻结表头
  sheet.setFrozenRows(1);
  
  // 自动调整列宽
  sheet.autoResizeColumns(1, headers.length);
}

// ========================================
// 📋 任务管理系统
// ========================================

/**
 * 设置任务管理表结构
 */
function 设置任务管理表() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('任务管理');
  
  // 清除现有数据
  sheet.clear();
  
  // 设置表头
  const headers = [
    '任务ID', '任务类型', '学生ID', '教师ID', '任务标题', '任务描述', '难度等级',
    '截止日期', '状态', '完成日期', '提交内容', '评分', '奖励积分', '反馈', '创建时间'
  ];
  
  const headerRange = sheet.getRange(1, 1, 1, headers.length);
  headerRange.setValues([headers]);
  
  // 美化表头
  headerRange.setFontWeight('bold');
  headerRange.setBackground('#9c27b0');
  headerRange.setFontColor('white');
  headerRange.setHorizontalAlignment('center');
  headerRange.setFontSize(11);
  
  // 设置数据验证
  const 任务类型选项 = ['日常任务', '技术作业', '英语练习', '项目任务', '考核任务', '额外挑战'];
  const 难度等级选项 = ['简单', '中等', '困难', '专家'];
  const 状态选项 = ['待开始', '进行中', '待审核', '已完成', '已过期'];
  
  // 任务类型下拉列表
  const typeRange = sheet.getRange(2, 2, 1000, 1);
  typeRange.setDataValidation(SpreadsheetApp.newDataValidation()
    .requireValueInList(任务类型选项, true)
    .build());
  
  // 难度等级下拉列表
  const difficultyRange = sheet.getRange(2, 7, 1000, 1);
  difficultyRange.setDataValidation(SpreadsheetApp.newDataValidation()
    .requireValueInList(难度等级选项, true)
    .build());
  
  // 状态下拉列表
  const statusRange = sheet.getRange(2, 9, 1000, 1);
  statusRange.setDataValidation(SpreadsheetApp.newDataValidation()
    .requireValueInList(状态选项, true)
    .build());
  
  // 设置格式
  sheet.getRange('H:H').setNumberFormat('yyyy-mm-dd');
  sheet.getRange('J:J').setNumberFormat('yyyy-mm-dd');
  sheet.getRange('L:L').setNumberFormat('0.0');
  sheet.getRange('M:M').setNumberFormat('#,##0');
  sheet.getRange('O:O').setNumberFormat('yyyy-mm-dd hh:mm:ss');
  
  // 冻结表头
  sheet.setFrozenRows(1);
  
  // 自动调整列宽
  sheet.autoResizeColumns(1, headers.length);
}

/**
 * 教师布置任务
 */
function 布置任务(教师ID, 学生ID, 任务类型, 任务标题, 任务描述, 难度等级, 截止日期) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('任务管理');
  const 今天 = new Date();
  const 任务ID = '任务' + Utilities.getUuid().substring(0, 8).toUpperCase();
  
  // 将截止日期字符串转换为Date对象
  const 截止日期对象 = typeof 截止日期 === 'string' ? new Date(截止日期) : 截止日期;
  
  // 根据难度等级计算奖励积分
  const 积分映射 = {
    '简单': 10,
    '中等': 20,
    '困难': 30,
    '专家': 50
  };
  const 奖励积分 = 积分映射[难度等级] || 10;
  
  const 新行数据 = [
    任务ID,           // 任务ID
    任务类型,          // 任务类型
    学生ID,           // 学生ID
    教师ID,           // 教师ID
    任务标题,          // 任务标题
    任务描述,          // 任务描述
    难度等级,          // 难度等级
    截止日期对象,       // 截止日期
    '待开始',          // 状态
    '',              // 完成日期
    '',              // 提交内容
    0,               // 评分
    奖励积分,          // 奖励积分
    '',              // 反馈
    今天              // 创建时间
  ];
  
  sheet.appendRow(新行数据);
  
  // 发送任务通知
  发送任务通知(学生ID, 任务标题, 截止日期);
  
  return 任务ID;
}

/**
 * 学生提交任务
 */
function 提交任务(任务ID, 提交内容) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('任务管理');
  const data = sheet.getDataRange().getValues();
  
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === 任务ID) {
      const 今天 = new Date();
      
      // 更新任务状态
      sheet.getRange(i + 1, 9).setValue('待审核');
      sheet.getRange(i + 1, 10).setValue(今天);
      sheet.getRange(i + 1, 11).setValue(提交内容);
      
      // 通知教师审核
      const 教师ID = data[i][3];
      通知教师审核(教师ID, 任务ID);
      
      break;
    }
  }
}

/**
 * 教师审核任务
 */
function 审核任务(任务ID, 评分, 反馈) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('任务管理');
  const data = sheet.getDataRange().getValues();
  
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === 任务ID) {
      const 学生ID = data[i][2];
      const 奖励积分 = data[i][12];
      
      // 更新任务状态
      sheet.getRange(i + 1, 9).setValue('已完成');
      sheet.getRange(i + 1, 12).setValue(评分);
      sheet.getRange(i + 1, 14).setValue(反馈);
      
      // 如果评分合格，给予积分奖励
      if (评分 >= 60) {
        添加学生积分(学生ID, 奖励积分, `完成任务: ${data[i][4]}`);
      }
      
      break;
    }
  }
}

// ========================================
// 🏆 奖励系统管理
// ========================================

/**
 * 设置奖励系统表结构
 */
function 设置奖励系统表() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('奖励系统');
  
  // 清除现有数据
  sheet.clear();
  
  // 设置表头
  const headers = [
    '记录ID', '学生ID', '积分类型', '积分数量', '原因描述', '相关任务ID', 
    '当前总积分', '徽章等级', '获得时间', '是否已兑换', '兑换内容'
  ];
  
  const headerRange = sheet.getRange(1, 1, 1, headers.length);
  headerRange.setValues([headers]);
  
  // 美化表头
  headerRange.setFontWeight('bold');
  headerRange.setBackground('#fbc02d');
  headerRange.setFontColor('white');
  headerRange.setHorizontalAlignment('center');
  headerRange.setFontSize(11);
  
  // 设置数据验证
  const 积分类型选项 = ['任务完成', '考试优秀', '项目展示', '帮助同学', '出勤奖励', '特殊贡献'];
  
  const typeRange = sheet.getRange(2, 3, 1000, 1);
  typeRange.setDataValidation(SpreadsheetApp.newDataValidation()
    .requireValueInList(积分类型选项, true)
    .build());
  
  // 设置格式
  sheet.getRange('D:D').setNumberFormat('#,##0');
  sheet.getRange('G:G').setNumberFormat('#,##0');
  sheet.getRange('I:I').setNumberFormat('yyyy-mm-dd hh:mm:ss');
  
  // 冻结表头
  sheet.setFrozenRows(1);
  
  // 自动调整列宽
  sheet.autoResizeColumns(1, headers.length);
  
  // 创建徽章等级设置
  创建徽章等级设置();
}

/**
 * 创建徽章等级设置
 */
function 创建徽章等级设置() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('奖励系统');
  
  // 在右侧添加徽章等级说明
  const 徽章说明 = [
    ['徽章等级', '所需积分', '徽章图标'],
    ['新人', '0', '🌱'],
    ['学习者', '100', '📚'],
    ['进步者', '300', '🚀'],
    ['优秀生', '600', '⭐'],
    ['学霸', '1000', '🏆'],
    ['大师', '1500', '👑'],
    ['传奇', '2500', '💎']
  ];
  
  const badgeRange = sheet.getRange(1, 13, 徽章说明.length, 3);
  badgeRange.setValues(徽章说明);
  
  // 美化徽章说明
  const badgeHeaderRange = sheet.getRange(1, 13, 1, 3);
  badgeHeaderRange.setFontWeight('bold');
  badgeHeaderRange.setBackground('#4caf50');
  badgeHeaderRange.setFontColor('white');
}

/**
 * 添加学生积分
 */
function 添加学生积分(学生ID, 积分数量, 原因描述, 相关任务ID = '') {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('奖励系统');
  const 学生Sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('学生信息');
  const 今天 = new Date();
  const 记录ID = '积分' + Utilities.getUuid().substring(0, 8).toUpperCase();
  
  // 获取学生当前总积分
  const 当前总积分 = 获取学生总积分(学生ID) + 积分数量;
  
  // 计算徽章等级
  const 徽章等级 = 计算徽章等级(当前总积分);
  
  const 新行数据 = [
    记录ID,           // 记录ID
    学生ID,           // 学生ID
    '任务完成',        // 积分类型
    积分数量,          // 积分数量
    原因描述,          // 原因描述
    相关任务ID,        // 相关任务ID
    当前总积分,        // 当前总积分
    徽章等级,          // 徽章等级
    今天,             // 获得时间
    false,           // 是否已兑换
    ''               // 兑换内容
  ];
  
  sheet.appendRow(新行数据);
  
  // 更新学生信息表中的总积分和徽章
  更新学生积分和徽章(学生ID, 当前总积分, 徽章等级);
  
  return 记录ID;
}

/**
 * 获取学生总积分
 */
function 获取学生总积分(学生ID) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('奖励系统');
  const data = sheet.getDataRange().getValues();
  
  let 总积分 = 0;
  for (let i = 1; i < data.length; i++) {
    if (data[i][1] === 学生ID) {
      总积分 += data[i][3] || 0;
    }
  }
  
  return 总积分;
}

/**
 * 计算徽章等级
 */
function 计算徽章等级(总积分) {
  if (总积分 >= 2500) return '传奇 💎';
  if (总积分 >= 1500) return '大师 👑';
  if (总积分 >= 1000) return '学霸 🏆';
  if (总积分 >= 600) return '优秀生 ⭐';
  if (总积分 >= 300) return '进步者 🚀';
  if (总积分 >= 100) return '学习者 📚';
  return '新人 🌱';
}

/**
 * 更新学生积分和徽章
 */
function 更新学生积分和徽章(学生ID, 总积分, 徽章等级) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('学生信息');
  const data = sheet.getDataRange().getValues();
  
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === 学生ID) {
      sheet.getRange(i + 1, 15).setValue(总积分);  // 总积分列
      sheet.getRange(i + 1, 16).setValue(徽章等级); // 当前徽章列
      break;
    }
  }
}

// ========================================
// 📊 现代化仪表板
// ========================================

/**
 * 设置仪表板
 */
function 设置仪表板() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('仪表板');
  sheet.clear();
  
  // 设置仪表板布局
  sheet.getRange('A1').setValue('🎓 IT前端训练营管理系统 V2.0');
  sheet.getRange('A1').setFontSize(28);
  sheet.getRange('A1').setFontWeight('bold');
  sheet.getRange('A1').setFontColor('#1a73e8');
  sheet.getRange('A1:H1').merge();
  sheet.getRange('A1').setHorizontalAlignment('center');
  
  // 副标题
  sheet.getRange('A2').setValue('🚀 智能化学习管理 | 任务系统 | 奖励机制');
  sheet.getRange('A2').setFontSize(14);
  sheet.getRange('A2').setFontColor('#5f6368');
  sheet.getRange('A2:H2').merge();
  sheet.getRange('A2').setHorizontalAlignment('center');
}

/**
 * 创建现代化仪表板
 */
function 创建现代化仪表板() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('仪表板');
  
  // 清除并重新设置
  sheet.clear();
  
  // 主标题
  const titleCell = sheet.getRange('A1');
  titleCell.setValue('🎓 IT前端训练营管理系统 V2.0');
  titleCell.setFontSize(32);
  titleCell.setFontWeight('bold');
  titleCell.setFontColor('#1a73e8');
  sheet.getRange('A1:J1').merge();
  titleCell.setHorizontalAlignment('center');
  
  // 副标题
  const subtitleCell = sheet.getRange('A2');
  subtitleCell.setValue('✨ 全新中文界面 | 智能任务系统 | 积分奖励机制 | 动态下拉菜单');
  subtitleCell.setFontSize(16);
  subtitleCell.setFontColor('#34a853');
  sheet.getRange('A2:J2').merge();
  subtitleCell.setHorizontalAlignment('center');
  
  // 创建统计卡片
  创建统计卡片();
  
  // 创建图表区域
  创建图表区域();
  
  // 创建最近活动
  创建最近活动();
  
  // 创建快捷操作
  创建快捷操作();
  
  // 自动调整列宽
  sheet.autoResizeColumns(1, 10);
}

/**
 * 创建统计卡片
 */
function 创建统计卡片() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('仪表板');
  
  // 卡片1: 学生总数
  sheet.getRange('A4').setValue('👥 学生总数');
  sheet.getRange('A5').setValue('=COUNTA(学生信息!A:A)-1');
  sheet.getRange('A4:B5').setBackground('#e3f2fd');
  sheet.getRange('A4:B5').setBorder(true, true, true, true, true, true);
  sheet.getRange('A4:B5').setFontWeight('bold');
  sheet.getRange('A5').setFontSize(20);
  
  // 卡片2: 在读学生
  sheet.getRange('C4').setValue('✅ 在读学生');
  sheet.getRange('C5').setValue('=COUNTIFS(学生信息!J:J,"在读")');
  sheet.getRange('C4:D5').setBackground('#e8f5e8');
  sheet.getRange('C4:D5').setBorder(true, true, true, true, true, true);
  sheet.getRange('C4:D5').setFontWeight('bold');
  sheet.getRange('C5').setFontSize(20);
  
  // 卡片3: 活跃任务
  sheet.getRange('E4').setValue('📋 活跃任务');
  sheet.getRange('E5').setValue('=COUNTIFS(任务管理!I:I,"进行中")+COUNTIFS(任务管理!I:I,"待审核")');
  sheet.getRange('E4:F5').setBackground('#fff3e0');
  sheet.getRange('E4:F5').setBorder(true, true, true, true, true, true);
  sheet.getRange('E4:F5').setFontWeight('bold');
  sheet.getRange('E5').setFontSize(20);
  
  // 卡片4: 总积分发放
  sheet.getRange('G4').setValue('🏆 总积分');
  sheet.getRange('G5').setValue('=SUM(奖励系统!D:D)');
  sheet.getRange('G4:H5').setBackground('#f3e5f5');
  sheet.getRange('G4:H5').setBorder(true, true, true, true, true, true);
  sheet.getRange('G4:H5').setFontWeight('bold');
  sheet.getRange('G5').setFontSize(20);
}

/**
 * 创建图表区域
 */
function 创建图表区域() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('仪表板');
  
  // 图表区域标题
  sheet.getRange('A7').setValue('📊 学习分析');
  sheet.getRange('A7').setFontSize(20);
  sheet.getRange('A7').setFontWeight('bold');
  sheet.getRange('A7').setFontColor('#1a73e8');
  
  // 学生等级分布
  sheet.getRange('A9').setValue('学生等级分布:');
  sheet.getRange('A9').setFontWeight('bold');
  sheet.getRange('A10').setValue('=QUERY(学生信息!A:F,"select F, count(A) where A is not null group by F label count(A) \'人数\'",1)');
  
  // 任务完成状态
  sheet.getRange('A13').setValue('任务状态分布:');
  sheet.getRange('A13').setFontWeight('bold');
  sheet.getRange('A14').setValue('=QUERY(任务管理!A:I,"select I, count(A) where A is not null group by I label count(A) \'数量\'",1)');
  
  // 徽章分布
  sheet.getRange('A17').setValue('徽章等级分布:');
  sheet.getRange('A17').setFontWeight('bold');
  sheet.getRange('A18').setValue('=QUERY(学生信息!A:P,"select P, count(A) where A is not null group by P label count(A) \'人数\'",1)');
}

/**
 * 创建最近活动
 */
function 创建最近活动() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('仪表板');
  
  // 最近活动标题
  sheet.getRange('F7').setValue('🕒 最近活动');
  sheet.getRange('F7').setFontSize(20);
  sheet.getRange('F7').setFontWeight('bold');
  sheet.getRange('F7').setFontColor('#1a73e8');
  
  // 最近新增学生
  sheet.getRange('F9').setValue('最近新增学生:');
  sheet.getRange('F9').setFontWeight('bold');
  sheet.getRange('F10').setValue('=QUERY(学生信息!A:R,"select A, B, R where A is not null order by R desc limit 5",1)');
  
  // 最近完成任务
  sheet.getRange('F16').setValue('最近完成任务:');
  sheet.getRange('F16').setFontWeight('bold');
  sheet.getRange('F17').setValue('=QUERY(任务管理!A:O,"select A, E, J where I = \'已完成\' order by J desc limit 5",1)');
}

/**
 * 创建快捷操作
 */
function 创建快捷操作() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('仪表板');
  
  // 快捷操作标题
  sheet.getRange('A25').setValue('⚡ 快捷操作');
  sheet.getRange('A25').setFontSize(20);
  sheet.getRange('A25').setFontWeight('bold');
  sheet.getRange('A25').setFontColor('#1a73e8');
  
  // 操作说明
  const 操作列表 = [
    '• 添加学生: 学生管理 > 添加新学生',
    '• 布置任务: 任务管理 > 新建任务',
    '• 查看积分: 奖励系统 > 积分记录',
    '• 生成报告: 报告分析 > 自定义报告',
    '• 系统设置: 系统设置 > 参数配置'
  ];
  
  for (let i = 0; i < 操作列表.length; i++) {
    sheet.getRange(27 + i, 1).setValue(操作列表[i]);
    sheet.getRange(27 + i, 1).setFontSize(12);
  }
}

/**
 * 更新仪表板
 */
function 更新仪表板() {
  // 强制重新计算所有公式
  SpreadsheetApp.flush();
  console.log('仪表板已更新');
}

// ========================================
// 🔧 辅助函数
// ========================================

/**
 * 发送任务通知（模拟）
 */
function 发送任务通知(学生ID, 任务标题, 截止日期) {
  console.log(`📧 发送任务通知给学生 ${学生ID}: ${任务标题}，截止日期: ${截止日期}`);
  // 这里可以集成邮件发送功能
}

/**
 * 通知教师审核（模拟）
 */
function 通知教师审核(教师ID, 任务ID) {
  console.log(`📧 通知教师 ${教师ID} 审核任务 ${任务ID}`);
  // 这里可以集成邮件发送功能
}

/**
 * 添加示例数据
 */
function 添加示例数据() {
  // 添加示例教师
  添加新教师('张老师', 'zhang@example.com', '13800138001', '前端开发', 5, 200);
  添加新教师('李老师', 'li@example.com', '13800138002', '英语教学', 3, 150);
  
  // 添加示例学生
  const 学生1 = 添加新学生('王小明', 'wang@example.com', '13900139001', new Date(), '初级', '张老师');
  const 学生2 = 添加新学生('刘小红', 'liu@example.com', '13900139002', new Date(), '中级', '李老师');
  
  // 添加示例任务
  布置任务('张老师', 学生1, '技术作业', '完成HTML基础练习', '完成HTML标签的使用练习，包括表格、表单、列表等', '简单', new Date(Date.now() + 7 * 24 * 60 * 60 * 1000));
  布置任务('李老师', 学生2, '英语练习', '英语口语练习', '录制5分钟的英语自我介绍视频', '中等', new Date(Date.now() + 5 * 24 * 60 * 60 * 1000));
}

// ========================================
// 🎯 顶部导航菜单
// ========================================

/**
 * 创建自定义菜单
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  
  // 创建主菜单
  ui.createMenu('🎓 训练营管理V2')
    .addSubMenu(ui.createMenu('👥 学生管理')
      .addItem('➕ 添加新学生', '显示添加学生对话框')
      .addItem('📊 学生统计', '显示学生统计')
      .addItem('🏆 积分排行', '显示积分排行榜'))
    
    .addSubMenu(ui.createMenu('👨‍🏫 教师管理')
      .addItem('➕ 添加新教师', '显示添加教师对话框')
      .addItem('📋 教师工作量', '显示教师工作量统计')
      .addItem('⭐ 教学评价', '显示教学评价'))
    
    .addSubMenu(ui.createMenu('📋 任务管理')
      .addItem('➕ 布置新任务', '显示布置任务界面')
      .addItem('📝 批量布置任务', '显示批量任务界面')
      .addItem('🔍 任务查询', '显示任务查询界面')
      .addItem('⏰ 任务提醒', '发送任务提醒'))
    
    .addSubMenu(ui.createMenu('🏆 奖励系统')
      .addItem('🎁 发放积分', '显示积分发放界面')
      .addItem('🏅 徽章管理', '显示徽章管理界面')
      .addItem('💰 积分兑换', '显示积分兑换界面')
      .addItem('📊 奖励统计', '显示奖励统计'))
    
    .addSubMenu(ui.createMenu('📊 报告分析')
      .addItem('📈 学习报告', '生成学习进度报告')
      .addItem('📋 任务报告', '生成任务完成报告')
      .addItem('🏆 积分报告', '生成积分分析报告')
      .addItem('💾 导出数据', '导出系统数据'))
    
    .addSeparator()
    .addItem('🔄 刷新仪表板', '更新仪表板')
    .addItem('⚙️ 系统设置', '显示系统设置')
    .addItem('🆘 帮助文档', '显示帮助信息')
    
    .addToUi();
  
  // 显示欢迎信息
  console.log('🎉 训练营管理系统V2已加载完成！');
}

// ========================================
// 📱 HTML界面对话框
// ========================================

/**
 * 显示添加学生对话框
 */
function 显示添加学生对话框() {
  const htmlOutput = HtmlService.createHtmlOutputFromFile('添加学生界面')
    .setWidth(600)
    .setHeight(500);
  
  SpreadsheetApp.getUi().showModalDialog(htmlOutput, '➕ 添加新学生');
}

/**
 * 显示添加教师对话框
 */
function 显示添加教师对话框() {
  const htmlOutput = HtmlService.createHtmlOutputFromFile('添加教师界面')
    .setWidth(600)
    .setHeight(500);
  
  SpreadsheetApp.getUi().showModalDialog(htmlOutput, '➕ 添加新教师');
}

/**
 * 显示布置任务界面
 */
function 显示布置任务界面() {
  const htmlOutput = HtmlService.createHtmlOutputFromFile('布置任务界面')
    .setWidth(700)
    .setHeight(600);
  
  SpreadsheetApp.getUi().showModalDialog(htmlOutput, '📋 布置新任务');
}

/**
 * 显示积分发放界面
 */
function 显示积分发放界面() {
  const htmlOutput = HtmlService.createHtmlOutputFromFile('积分管理界面')
    .setWidth(900)
    .setHeight(700);
  
  SpreadsheetApp.getUi().showModalDialog(htmlOutput, '🏆 积分奖励管理');
}

/**
 * 显示批量任务界面
 */
function 显示批量任务界面() {
  const ui = SpreadsheetApp.getUi();
  
  const result = ui.prompt(
    '📝 批量布置任务',
    '请输入任务信息，格式：任务类型|任务标题|任务描述|难度等级|截止天数\n例如：技术作业|JavaScript练习|完成基础练习|简单|7',
    ui.ButtonSet.OK_CANCEL
  );
  
  if (result.getSelectedButton() === ui.Button.OK) {
    const input = result.getResponseText();
    const parts = input.split('|');
    
    if (parts.length >= 5) {
      const 任务类型 = parts[0].trim();
      const 任务标题 = parts[1].trim();
      const 任务描述 = parts[2].trim();
      const 难度等级 = parts[3].trim();
      const 截止天数 = parseInt(parts[4].trim());
      
      // 获取所有在读学生
      const 学生列表 = 获取学生列表().filter(s => s.status === '在读');
      const 教师列表 = 获取教师列表();
      
      if (学生列表.length === 0) {
        ui.alert('❌ 没有找到在读学生');
        return;
      }
      
      if (教师列表.length === 0) {
        ui.alert('❌ 没有找到可用教师');
        return;
      }
      
      const 教师ID = 教师列表[0].id; // 使用第一个教师
      const 截止日期 = new Date(Date.now() + 截止天数 * 24 * 60 * 60 * 1000);
      
      let 成功计数 = 0;
      
      学生列表.forEach(学生 => {
        try {
          布置任务(教师ID, 学生.id, 任务类型, 任务标题, 任务描述, 难度等级, 截止日期);
          成功计数++;
        } catch (error) {
          console.error(`为学生 ${学生.name} 布置任务失败:`, error);
        }
      });
      
      ui.alert('✅ 批量任务布置完成', `成功为 ${成功计数} 名学生布置了任务`, ui.ButtonSet.OK);
    } else {
      ui.alert('❌ 输入格式错误', '请按照指定格式输入任务信息', ui.ButtonSet.OK);
    }
  }
}

/**
 * 显示任务查询界面
 */
function 显示任务查询界面() {
  const ui = SpreadsheetApp.getUi();
  
  const result = ui.prompt(
    '🔍 任务查询',
    '请选择查询方式：\n1 - 按学生查询\n2 - 按状态查询\n3 - 按截止日期查询',
    ui.ButtonSet.OK_CANCEL
  );
  
  if (result.getSelectedButton() === ui.Button.OK) {
    const 查询方式 = result.getResponseText();
    
    let 查询结果 = '';
    
    switch (查询方式) {
      case '1':
        查询结果 = 按学生查询任务();
        break;
      case '2':
        查询结果 = 按状态查询任务();
        break;
      case '3':
        查询结果 = 按日期查询任务();
        break;
      default:
        ui.alert('❌ 无效的查询方式');
        return;
    }
    
    ui.alert('🔍 查询结果', 查询结果, ui.ButtonSet.OK);
  }
}

/**
 * 显示徽章管理界面
 */
function 显示徽章管理界面() {
  const ui = SpreadsheetApp.getUi();
  
  const 徽章统计 = 获取徽章统计();
  
  let 统计信息 = `🏅 徽章管理\n\n`;
  统计信息 += `📊 徽章分布统计：\n`;
  
  Object.entries(徽章统计).forEach(([徽章, 人数]) => {
    统计信息 += `${徽章}: ${人数}人\n`;
  });
  
  ui.alert('徽章管理', 统计信息, ui.ButtonSet.OK);
}

/**
 * 显示积分兑换界面
 */
function 显示积分兑换界面() {
  const ui = SpreadsheetApp.getUi();
  
  const 兑换说明 = `💰 积分兑换\n\n` +
                  `🎁 可兑换奖励：\n` +
                  `• 100积分 - 学习资料包\n` +
                  `• 200积分 - 1对1辅导1小时\n` +
                  `• 500积分 - 项目指导\n` +
                  `• 1000积分 - 就业推荐\n\n` +
                  `如需兑换请联系管理员`;
  
  ui.alert('积分兑换', 兑换说明, ui.ButtonSet.OK);
}

/**
 * 显示奖励统计
 */
function 显示奖励统计() {
  const 奖励Sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('奖励系统');
  const 数据 = 奖励Sheet.getDataRange().getValues();
  
  let 总积分 = 0;
  let 发放次数 = 0;
  const 类型统计 = {};
  
  for (let i = 1; i < 数据.length; i++) {
    if (数据[i][0]) { // 如果记录ID不为空
      总积分 += 数据[i][3] || 0;
      发放次数++;
      
      const 类型 = 数据[i][2];
      类型统计[类型] = (类型统计[类型] || 0) + 1;
    }
  }
  
  let 统计信息 = `📊 奖励统计\n\n`;
  统计信息 += `🏆 总积分发放: ${总积分}分\n`;
  统计信息 += `📈 发放次数: ${发放次数}次\n`;
  统计信息 += `📊 平均每次: ${发放次数 > 0 ? (总积分 / 发放次数).toFixed(1) : 0}分\n\n`;
  统计信息 += `📋 积分类型分布:\n`;
  
  Object.entries(类型统计).forEach(([类型, 次数]) => {
    统计信息 += `• ${类型}: ${次数}次\n`;
  });
  
  SpreadsheetApp.getUi().alert('奖励统计', 统计信息, SpreadsheetApp.getUi().ButtonSet.OK);
}

/**
 * 显示教学评价
 */
function 显示教学评价() {
  const ui = SpreadsheetApp.getUi();
  
  const 教师Sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('教师信息');
  const 数据 = 教师Sheet.getDataRange().getValues();
  
  let 评价信息 = `⭐ 教学评价\n\n`;
  
  for (let i = 1; i < 数据.length; i++) {
    if (数据[i][0]) { // 如果教师ID不为空
      const 姓名 = 数据[i][1];
      const 专业领域 = 数据[i][4];
      const 教学评分 = 数据[i][10] || 5.0;
      
      评价信息 += `👨‍🏫 ${姓名} (${专业领域}): ${教学评分}分\n`;
    }
  }
  
  ui.alert('教学评价', 评价信息, ui.ButtonSet.OK);
}

/**
 * 发送任务提醒
 */
function 发送任务提醒() {
  const 提醒数量 = 批量发送任务提醒();
  
  SpreadsheetApp.getUi().alert(
    '⏰ 任务提醒',
    `已发送 ${提醒数量} 条任务提醒`,
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}

/**
 * 生成学习进度报告
 */
function 生成学习进度报告() {
  const ui = SpreadsheetApp.getUi();
  
  const 学生Sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('学生信息');
  const 技术Sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('技术学习');
  const 英语Sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('英语学习');
  
  const 学生数据 = 学生Sheet.getDataRange().getValues();
  const 技术数据 = 技术Sheet.getDataRange().getValues();
  const 英语数据 = 英语Sheet.getDataRange().getValues();
  
  let 总学生数 = 0;
  let 在读学生数 = 0;
  let 总技术课程 = 0;
  let 完成技术课程 = 0;
  let 总英语课程 = 0;
  let 完成英语课程 = 0;
  
  // 统计学生数据
  for (let i = 1; i < 学生数据.length; i++) {
    if (学生数据[i][0]) {
      总学生数++;
      if (学生数据[i][9] === '在读') {
        在读学生数++;
      }
    }
  }
  
  // 统计技术学习
  for (let i = 1; i < 技术数据.length; i++) {
    if (技术数据[i][0]) {
      总技术课程++;
      if (技术数据[i][4] === '已完成') {
        完成技术课程++;
      }
    }
  }
  
  // 统计英语学习
  for (let i = 1; i < 英语数据.length; i++) {
    if (英语数据[i][0]) {
      总英语课程++;
      if (英语数据[i][4] === '已完成') {
        完成英语课程++;
      }
    }
  }
  
  const 技术完成率 = 总技术课程 > 0 ? (完成技术课程 / 总技术课程 * 100).toFixed(1) : 0;
  const 英语完成率 = 总英语课程 > 0 ? (完成英语课程 / 总英语课程 * 100).toFixed(1) : 0;
  
  const 报告内容 = `📈 学习进度报告\n\n` +
                 `👥 学生概况:\n` +
                 `• 总学生数: ${总学生数}人\n` +
                 `• 在读学生: ${在读学生数}人\n` +
                 `• 在读率: ${总学生数 > 0 ? (在读学生数 / 总学生数 * 100).toFixed(1) : 0}%\n\n` +
                 `💻 技术学习:\n` +
                 `• 总课程数: ${总技术课程}\n` +
                 `• 已完成: ${完成技术课程}\n` +
                 `• 完成率: ${技术完成率}%\n\n` +
                 `🗣️ 英语学习:\n` +
                 `• 总课程数: ${总英语课程}\n` +
                 `• 已完成: ${完成英语课程}\n` +
                 `• 完成率: ${英语完成率}%\n\n` +
                 `📊 综合完成率: ${((parseFloat(技术完成率) + parseFloat(英语完成率)) / 2).toFixed(1)}%`;
  
  ui.alert('学习进度报告', 报告内容, ui.ButtonSet.OK);
}

/**
 * 生成任务完成报告
 */
function 生成任务完成报告() {
  const ui = SpreadsheetApp.getUi();
  
  const 任务统计 = 获取任务统计();
  
  const 报告内容 = `📋 任务完成报告\n\n` +
                 `📊 任务概况:\n` +
                 `• 总任务数: ${任务统计.总任务数}\n` +
                 `• 已完成: ${任务统计.已完成}\n` +
                 `• 进行中: ${任务统计.进行中}\n` +
                 `• 待审核: ${任务统计.待审核}\n` +
                 `• 待开始: ${任务统计.待开始}\n` +
                 `• 已过期: ${任务统计.已过期}\n\n` +
                 `📈 完成率: ${任务统计.完成率}%\n\n` +
                 `💡 建议:\n` +
                 `${任务统计.完成率 < 60 ? '• 完成率偏低，建议加强督促\n' : ''}` +
                 `${任务统计.已过期 > 0 ? `• 有${任务统计.已过期}个过期任务需要处理\n` : ''}` +
                 `${任务统计.待审核 > 5 ? '• 待审核任务较多，建议及时处理\n' : ''}`;
  
  ui.alert('任务完成报告', 报告内容, ui.ButtonSet.OK);
}

/**
 * 生成积分分析报告
 */
function 生成积分分析报告() {
  const ui = SpreadsheetApp.getUi();
  
  const 积分排行 = 获取学生积分排行();
  const 积分历史 = 获取积分历史();
  
  let 总积分 = 0;
  let 平均积分 = 0;
  let 最高积分 = 0;
  let 最低积分 = Number.MAX_VALUE;
  
  积分排行.forEach(学生 => {
    总积分 += 学生.totalPoints;
    最高积分 = Math.max(最高积分, 学生.totalPoints);
    最低积分 = Math.min(最低积分, 学生.totalPoints);
  });
  
  if (积分排行.length > 0) {
    平均积分 = (总积分 / 积分排行.length).toFixed(1);
    if (最低积分 === Number.MAX_VALUE) 最低积分 = 0;
  }
  
  // 统计积分类型
  const 类型统计 = {};
  积分历史.forEach(记录 => {
    类型统计[记录.type] = (类型统计[记录.type] || 0) + 记录.points;
  });
  
  let 类型分析 = '';
  Object.entries(类型统计).forEach(([类型, 积分]) => {
    类型分析 += `• ${类型}: ${积分}分\n`;
  });
  
  const 报告内容 = `🏆 积分分析报告\n\n` +
                 `📊 积分概况:\n` +
                 `• 总积分发放: ${总积分}分\n` +
                 `• 平均积分: ${平均积分}分\n` +
                 `• 最高积分: ${最高积分}分\n` +
                 `• 最低积分: ${最低积分}分\n` +
                 `• 参与学生: ${积分排行.length}人\n\n` +
                 `📈 积分类型分布:\n${类型分析}\n` +
                 `🏅 前三名:\n` +
                 `${积分排行.length > 0 ? `🥇 ${积分排行[0].name}: ${积分排行[0].totalPoints}分\n` : ''}` +
                 `${积分排行.length > 1 ? `🥈 ${积分排行[1].name}: ${积分排行[1].totalPoints}分\n` : ''}` +
                 `${积分排行.length > 2 ? `🥉 ${积分排行[2].name}: ${积分排行[2].totalPoints}分\n` : ''}`;
  
  ui.alert('积分分析报告', 报告内容, ui.ButtonSet.OK);
}

/**
 * 获取所有教师列表（供下拉菜单使用）
 */
function 获取教师列表() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('教师信息');
  const data = sheet.getDataRange().getValues();
  
  const 教师列表 = [];
  for (let i = 1; i < data.length; i++) {
    if (data[i][0]) { // 如果教师ID不为空
      教师列表.push({
        id: data[i][0],
        name: data[i][1],
        specialization: data[i][4]
      });
    }
  }
  
  return 教师列表;
}

/**
 * 获取所有学生列表（供下拉菜单使用）
 */
function 获取学生列表() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('学生信息');
  const data = sheet.getDataRange().getValues();
  
  const 学生列表 = [];
  for (let i = 1; i < data.length; i++) {
    if (data[i][0]) { // 如果学生ID不为空
      学生列表.push({
        id: data[i][0],
        name: data[i][1],
        level: data[i][5],
        status: data[i][9]
      });
    }
  }
  
  return 学生列表;
}

// ========================================
// 📊 其他工作表设置
// ========================================

/**
 * 设置英语学习表
 */
function 设置英语学习表() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('英语学习');
  
  sheet.clear();
  
  const headers = [
    '学习ID', '学生ID', '技能类型', '等级', '状态', '进度百分比', '分数',
    '开始日期', '完成日期', '指导老师', '学习资源', '备注'
  ];
  
  const headerRange = sheet.getRange(1, 1, 1, headers.length);
  headerRange.setValues([headers]);
  headerRange.setFontWeight('bold');
  headerRange.setBackground('#4caf50');
  headerRange.setFontColor('white');
  headerRange.setHorizontalAlignment('center');
  
  sheet.autoResizeColumns(1, headers.length);
}

/**
 * 设置求职管理表
 */
function 设置求职管理表() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('求职管理');
  
  sheet.clear();
  
  const headers = [
    '申请ID', '学生ID', '公司名称', '职位名称', '状态', '申请日期',
    '面试日期', '回复状态', '薪资范围', '工作地点', '备注'
  ];
  
  const headerRange = sheet.getRange(1, 1, 1, headers.length);
  headerRange.setValues([headers]);
  headerRange.setFontWeight('bold');
  headerRange.setBackground('#e91e63');
  headerRange.setFontColor('white');
  headerRange.setHorizontalAlignment('center');
  
  sheet.autoResizeColumns(1, headers.length);
}

/**
 * 设置项目作品表
 */
function 设置项目作品表() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('项目作品');
  
  sheet.clear();
  
  const headers = [
    '项目ID', '学生ID', '项目名称', '项目类型', '状态', '进度百分比',
    '开始日期', '预期完成日期', '实际完成日期', '技术栈', 'GitHub链接',
    '演示链接', '评分', '反馈', '备注'
  ];
  
  const headerRange = sheet.getRange(1, 1, 1, headers.length);
  headerRange.setValues([headers]);
  headerRange.setFontWeight('bold');
  headerRange.setBackground('#795548');
  headerRange.setFontColor('white');
  headerRange.setHorizontalAlignment('center');
  
  sheet.autoResizeColumns(1, headers.length);
}

/**
 * 设置考核评估表
 */
function 设置考核评估表() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('考核评估');
  
  sheet.clear();
  
  const headers = [
    '考核ID', '学生ID', '考核类型', '考核科目', '分数', '满分',
    '考核日期', '监考老师', '状态', '反馈意见', '重考日期', '备注'
  ];
  
  const headerRange = sheet.getRange(1, 1, 1, headers.length);
  headerRange.setValues([headers]);
  headerRange.setFontWeight('bold');
  headerRange.setBackground('#607d8b');
  headerRange.setFontColor('white');
  headerRange.setHorizontalAlignment('center');
  
  sheet.autoResizeColumns(1, headers.length);
}

/**
 * 设置报告分析表
 */
function 设置报告分析表() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('报告分析');
  
  sheet.clear();
  
  const headers = [
    '报告类型', '时间范围', '学生总数', '在读学生', '毕业率', '就业率',
    '平均技术分数', '平均英语分数', '任务完成率', '积分发放总数'
  ];
  
  const headerRange = sheet.getRange(1, 1, 1, headers.length);
  headerRange.setValues([headers]);
  headerRange.setFontWeight('bold');
  headerRange.setBackground('#ff9800');
  headerRange.setFontColor('white');
  headerRange.setHorizontalAlignment('center');
  
  sheet.autoResizeColumns(1, headers.length);
}

/**
 * 设置系统设置表
 */
function 设置系统设置表() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('系统设置');
  
  sheet.clear();
  
  const 设置项 = [
    ['设置项', '设置值', '说明'],
    ['训练营名称', 'IT前端训练营V2', '训练营的名称'],
    ['学习模式', '线上+线下', '学习交付方式'],
    ['课程周期', '6个月', '完整课程周期'],
    ['每班最大人数', '30', '每个班级的最大学生数'],
    ['及格分数', '60', '考核及格最低分数'],
    ['就业目标', '85%', '目标就业率'],
    ['英语要求', 'B1', '要求的英语水平'],
    ['技术栈', 'HTML/CSS, JavaScript, React, Node.js', '核心技术栈'],
    ['考核频率', '每周', '考核进行频率'],
    ['师生比例', '1:15', '教师学生比例'],
    ['积分制度', '启用', '是否启用积分奖励制度'],
    ['徽章系统', '启用', '是否启用徽章系统']
  ];
  
  const dataRange = sheet.getRange(1, 1, 设置项.length, 3);
  dataRange.setValues(设置项);
  
  // 格式化表头
  const headerRange = sheet.getRange(1, 1, 1, 3);
  headerRange.setFontWeight('bold');
  headerRange.setBackground('#3f51b5');
  headerRange.setFontColor('white');
  headerRange.setHorizontalAlignment('center');
  
  sheet.autoResizeColumns(1, 3);
}

// ========================================
// 🎯 最终初始化函数
// ========================================

/**
 * 手动运行初始化（用于测试）
 */
function 手动初始化() {
  初始化训练营系统V2();
}
