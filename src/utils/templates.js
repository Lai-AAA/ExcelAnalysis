// 活动/讲座类模板配置
export const activityTemplate = {
  type: 'activity',
  columns: [
    { label: '序号', field: '序号', width: 4.75 },
    { label: '班级', field: '行政班级', width: 27 },
    { label: '姓名', field: '姓名', width: 8.25 }
  ],
  rowHeights: {
    title: 69,
    note: 24.5,
    header: 20,
    body: 20
  },
  fontSizes: {
    title: 18,
    note: 12,
    body: 11
  },
  generateTitle: (semester, activityName) => {
    return `${semester}工学院${activityName}加分名单`
  },
  generateNote: (scoreType, score) => {
    return `注：以下同学每人加${score}${scoreType}分`
  }
}

// 比赛类模板配置
export const competitionTemplate = {
  type: 'competition',
  columns: [
    { label: '序号', field: '序号', width: 4.75 },
    { label: '班级', field: '行政班级', width: 27 },
    { label: '姓名', field: '姓名', width: 8.25 },
    { label: '奖项', field: '奖项', width: 16.25 },
    { label: '加分数', field: '加分分数', width: 7.25 }
  ],
  rowHeights: {
    title: 69,
    note: 24.5,
    header: 20,
    body: 20
  },
  fontSizes: {
    title: 18,
    note: 12,
    body: 11
  },
  generateTitle: (semester, activityName) => {
    return `${semester}工学院${activityName}加分名单`
  },
  generateNote: (scoreType) => {
    return `注：以下同学加${scoreType}，具体分数如下`
  }
}

/**
 * 根据活动类型匹配模板
 */
export function getTemplateByActivityType(activityType) {
  if (activityType === '活动' || activityType === '讲座') {
    return activityTemplate
  } else if (activityType === '比赛') {
    return competitionTemplate
  }
  return activityTemplate // 默认返回活动模板
}

/**
 * 默认加分规则
 */
export const defaultScoreRules = {
  '一等奖': 4,
  '二等奖': 3,
  '三等奖': 2,
  '优秀奖': 1,
  '参与奖': 0.5
}

