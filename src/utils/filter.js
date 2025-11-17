/**
 * 多维度数据筛选
 */
export function filterData(data, filters) {
  let result = [...data]
  
  // 学年学期筛选
  if (filters.semester) {
    result = result.filter(row => row['学年学期'] === filters.semester)
  }
  
  // 活动类型筛选
  if (filters.activityTypes && filters.activityTypes.length > 0) {
    result = result.filter(row => filters.activityTypes.includes(row['活动类型']))
  }
  
  // 活动名称筛选（模糊搜索）
  if (filters.activityName) {
    const keyword = filters.activityName.toLowerCase()
    result = result.filter(row => {
      const name = String(row['活动名称'] || '').toLowerCase()
      return name.includes(keyword)
    })
  }
  
  // 加分类型筛选
  if (filters.scoreType) {
    result = result.filter(row => row['加分类型'] === filters.scoreType)
  }
  
  // 加分分数范围筛选
  if (filters.minScore !== undefined && filters.minScore !== null) {
    result = result.filter(row => row['加分分数'] >= filters.minScore)
  }
  if (filters.maxScore !== undefined && filters.maxScore !== null) {
    result = result.filter(row => row['加分分数'] <= filters.maxScore)
  }
  
  // 班级筛选
  if (filters.classes && filters.classes.length > 0) {
    result = result.filter(row => filters.classes.includes(row['行政班级']))
  }
  
  // 是否包含其他学院
  if (filters.includeOtherCollege === false) {
    result = result.filter(row => row['是否为工学院举办'] === '是')
  }
  
  // 排序
  if (filters.sortByClass !== false) {
    result.sort((a, b) => {
      const classA = String(a['行政班级'] || '')
      const classB = String(b['行政班级'] || '')
      if (classA !== classB) {
        return classA.localeCompare(classB, 'zh-CN')
      }
      const nameA = String(a['姓名'] || '')
      const nameB = String(b['姓名'] || '')
      return nameA.localeCompare(nameB, 'zh-CN')
    })
  }
  
  return result
}

/**
 * 获取所有唯一的学年学期
 */
export function getUniqueSemesters(data) {
  const semesters = new Set()
  data.forEach(row => {
    if (row['学年学期']) {
      semesters.add(row['学年学期'])
    }
  })
  return Array.from(semesters).sort()
}

/**
 * 获取所有唯一的活动类型
 */
export function getUniqueActivityTypes(data) {
  const types = new Set()
  data.forEach(row => {
    if (row['活动类型']) {
      types.add(row['活动类型'])
    }
  })
  return Array.from(types).sort()
}

/**
 * 获取所有唯一的班级
 */
export function getUniqueClasses(data) {
  const classes = new Set()
  data.forEach(row => {
    if (row['行政班级']) {
      classes.add(row['行政班级'])
    }
  })
  return Array.from(classes).sort()
}

