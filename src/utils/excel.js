import * as XLSX from 'xlsx'
import ExcelJS from 'exceljs'

// 必需的15个核心字段
const REQUIRED_FIELDS = [
  '序号', '学年学期', '是否为工学院举办', '行政班级', '学号', 
  '姓名', '活动类型', '活动名称', '加分类型', '加分分数', 
  '奖项', '部门', '负责人', '联系电话', '备注'
]

/**
 * 解析Excel文件
 * @param {File} file - Excel文件对象
 * @returns {Promise<Object>} 解析结果
 */
export async function parseExcelFile(file) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader()
    
    reader.onload = (e) => {
      try {
        const data = new Uint8Array(e.target.result)
        const workbook = XLSX.read(data, { type: 'array' })
        
        // 检查是否存在"数据源"工作表
        if (!workbook.SheetNames.includes('数据源')) {
          reject(new Error('请上传包含"数据源"工作表的标准文件'))
          return
        }
        
        const worksheet = workbook.Sheets['数据源']
        const jsonData = XLSX.utils.sheet_to_json(worksheet, { defval: '' })
        
        // 验证必需字段
        const missingFields = REQUIRED_FIELDS.filter(field => !jsonData[0] || !(field in jsonData[0]))
        if (missingFields.length > 0) {
          reject(new Error(`缺少以下列：${missingFields.join('、')}，请检查数据源格式`))
          return
        }
        
        // 数据处理
        const processedData = processData(jsonData)
        
        resolve({
          rawData: jsonData,
          processedData: processedData.valid,
          invalidData: processedData.invalid,
          stats: processedData.stats
        })
      } catch (error) {
        reject(new Error(`文件解析失败：${error.message}`))
      }
    }
    
    reader.onerror = () => {
      reject(new Error('文件读取失败，请检查文件是否损坏'))
    }
    
    reader.readAsArrayBuffer(file)
  })
}

/**
 * 处理数据：去重、格式转换、有效性校验
 */
function processData(data) {
  const valid = []
  const invalid = []
  const seen = new Set()
  const classes = new Set()
  const activityTypes = new Set()
  
  data.forEach((row, index) => {
    // 去除首尾空格
    const name = String(row['姓名'] || '').trim()
    const activityName = String(row['活动名称'] || '').trim()
    const className = String(row['行政班级'] || '').trim()
    
    // 有效性校验
    if (!name || !activityName) {
      invalid.push({ ...row, reason: '姓名为空或活动名称为空', rowIndex: index + 2 })
      return
    }
    
    // 去重：按"班级+姓名+活动名称"
    const key = `${className}|${name}|${activityName}`
    if (seen.has(key)) {
      invalid.push({ ...row, reason: '重复数据', rowIndex: index + 2 })
      return
    }
    seen.add(key)
    
    // 格式转换
    const processedRow = {
      ...row,
      姓名: name,
      活动名称: activityName,
      行政班级: className,
      加分分数: parseFloat(row['加分分数']) || 0,
      学号: String(row['学号'] || '').trim()
    }
    
    valid.push(processedRow)
    classes.add(className)
    activityTypes.add(row['活动类型'])
  })
  
  return {
    valid,
    invalid,
    stats: {
      total: data.length,
      valid: valid.length,
      invalid: invalid.length,
      classes: classes.size,
      activityTypes: activityTypes.size
    }
  }
}

/**
 * 验证考勤状态
 * @param {Array} data - 数据数组
 * @param {Object} options - 考勤验证选项
 * @returns {Object} 已考勤和未考勤数据
 */
export function validateAttendance(data, options = {}) {
  const { keywords = ['已考勤', '到场', '参与'], attendanceFile = null } = options
  const attended = []
  const unattended = []
  
  // 缺勤关键词列表
  const absenceKeywords = ['缺勤', '迟到', '早退', '请假']
  
  data.forEach(row => {
    const remark = String(row['备注'] || '').trim()
    
    // 首先检查是否为缺勤情况
    const isAbsent = absenceKeywords.some(keyword => remark.includes(keyword))
    if (isAbsent) {
      unattended.push({
        ...row,
        未考勤原因: `备注中包含缺勤信息：${remark}`
      })
      return
    }
    
    // 检查是否为全勤（可加分）
    if (remark === '全勤') {
      attended.push(row)
      return
    }
    
    // 其他情况：检查是否包含考勤关键词
    const isAttended = keywords.some(keyword => remark.includes(keyword))
    
    if (isAttended) {
      attended.push(row)
    } else {
      unattended.push({
        ...row,
        未考勤原因: '备注中未找到考勤关键词'
      })
    }
  })
  
  return { attended, unattended }
}

/**
 * 排序数据：按班级排序，其他学院放最后
 * @param {Array} data - 数据数组
 * @returns {Array} 排序后的数据
 */
function sortDataByClass(data) {
  // 工学院班级关键词
  const engineeringClasses = [
    '网络工程', '软件工程', '软件工程卓越工程师班', '通信工程',
    '食品', '食品超越班', '食品营养', '数字媒体本', '数字媒体创意班',
    '大数据', '电子信息工程', '人工智能', '人工智能物联网班'
  ]
  
  // 判断是否为工学院
  const isEngineeringCollege = (className) => {
    if (!className) return false
    return engineeringClasses.some(cls => className.includes(cls))
  }
  
  // 分离工学院和其他学院
  const engineeringData = []
  const otherCollegeData = []
  
  data.forEach(row => {
    const className = String(row['行政班级'] || '').trim()
    if (isEngineeringCollege(className)) {
      engineeringData.push(row)
    } else {
      otherCollegeData.push(row)
    }
  })
  
  // 对工学院数据按班级和姓名排序
  engineeringData.sort((a, b) => {
    const classA = String(a['行政班级'] || '')
    const classB = String(b['行政班级'] || '')
    if (classA !== classB) {
      return classA.localeCompare(classB, 'zh-CN')
    }
    const nameA = String(a['姓名'] || '')
    const nameB = String(b['姓名'] || '')
    return nameA.localeCompare(nameB, 'zh-CN')
  })
  
  // 合并：工学院在前，其他学院在后
  return [...engineeringData, ...otherCollegeData]
}

/**
 * 排序数据：按奖项排序（比赛类型）
 * @param {Array} data - 数据数组
 * @returns {Array} 排序后的数据
 */
function sortDataByAward(data) {
  // 奖项优先级顺序
  const awardOrder = {
    '一等奖': 1,
    '二等奖': 2,
    '三等奖': 3,
    '优秀奖': 4,
    '参与奖': 5
  }
  
  // 获取奖项优先级
  const getAwardPriority = (award) => {
    const awardStr = String(award || '').trim()
    // 检查是否包含奖项关键词
    for (const [key, priority] of Object.entries(awardOrder)) {
      if (awardStr.includes(key)) {
        return priority
      }
    }
    // 如果没有匹配的奖项，放在最后
    return 999
  }
  
  // 工学院班级关键词
  const engineeringClasses = [
    '网络工程', '软件工程', '软件工程卓越工程师班', '通信工程',
    '食品', '食品超越班', '食品营养', '数字媒体本', '数字媒体创意班',
    '大数据', '电子信息工程', '人工智能', '人工智能物联网班'
  ]
  
  // 判断是否为工学院
  const isEngineeringCollege = (className) => {
    if (!className) return false
    return engineeringClasses.some(cls => className.includes(cls))
  }
  
  // 分离工学院和其他学院
  const engineeringData = []
  const otherCollegeData = []
  
  data.forEach(row => {
    const className = String(row['行政班级'] || '').trim()
    if (isEngineeringCollege(className)) {
      engineeringData.push(row)
    } else {
      otherCollegeData.push(row)
    }
  })
  
  // 对工学院数据按奖项、班级和姓名排序
  engineeringData.sort((a, b) => {
    const awardA = getAwardPriority(a['奖项'])
    const awardB = getAwardPriority(b['奖项'])
    
    // 首先按奖项排序
    if (awardA !== awardB) {
      return awardA - awardB
    }
    
    // 奖项相同，按班级排序
    const classA = String(a['行政班级'] || '')
    const classB = String(b['行政班级'] || '')
    if (classA !== classB) {
      return classA.localeCompare(classB, 'zh-CN')
    }
    
    // 班级相同，按姓名排序
    const nameA = String(a['姓名'] || '')
    const nameB = String(b['姓名'] || '')
    return nameA.localeCompare(nameB, 'zh-CN')
  })
  
  // 对其他学院数据也按奖项排序
  otherCollegeData.sort((a, b) => {
    const awardA = getAwardPriority(a['奖项'])
    const awardB = getAwardPriority(b['奖项'])
    
    if (awardA !== awardB) {
      return awardA - awardB
    }
    
    const classA = String(a['行政班级'] || '')
    const classB = String(b['行政班级'] || '')
    if (classA !== classB) {
      return classA.localeCompare(classB, 'zh-CN')
    }
    
    const nameA = String(a['姓名'] || '')
    const nameB = String(b['姓名'] || '')
    return nameA.localeCompare(nameB, 'zh-CN')
  })
  
  // 合并：工学院在前，其他学院在后
  return [...engineeringData, ...otherCollegeData]
}

/**
 * 导出Excel文件
 * @param {Array} data - 数据数组
 * @param {Object} template - 模板配置
 * @param {String} filename - 文件名
 */
export async function exportExcel(data, template, filename) {
  // 判断是否为比赛类型
  const isCompetition = template.type === 'competition'
  
  // 根据类型选择排序方式：比赛类型按奖项排序，其他按班级排序
  const sortedData = isCompetition ? sortDataByAward(data) : sortDataByClass(data)
  
  // 创建Excel工作簿
  const workbook = new ExcelJS.Workbook()
  const worksheet = workbook.addWorksheet('加分名单')
  
  // 定义样式
  const titleStyle = {
    font: { name: '微软雅黑', size: template.fontSizes?.title || 18, bold: true },
    alignment: { horizontal: 'center', vertical: 'middle', wrapText: true }
  }
  
  const noteStyle = {
    font: { name: '微软雅黑', size: template.fontSizes?.note || 12, color: { argb: 'FFFF0000' } }, // 红色
    alignment: { horizontal: 'center', vertical: 'middle', wrapText: true }
  }
  
  const headerStyle = {
    font: { name: '微软雅黑', size: template.fontSizes?.body || 11, bold: true },
    alignment: { horizontal: 'center', vertical: 'middle' }
  }
  
  const bodyStyle = {
    font: { name: '微软雅黑', size: template.fontSizes?.body || 11 },
    alignment: { horizontal: 'center', vertical: 'middle' }
  }
  
  let currentRow = 1
  
  if (isCompetition) {
    // 比赛类型：单列布局（A、B、C、D、E列）
    // 设置列宽
    worksheet.getColumn(1).width = 4.75  // A列：序号
    worksheet.getColumn(2).width = 27   // B列：行政班级
    worksheet.getColumn(3).width = 8.25 // C列：姓名
    worksheet.getColumn(4).width = 16.25 // D列：奖项
    worksheet.getColumn(5).width = 7.25  // E列：加分数
    
    const totalCols = 5
    
    // 生成标题
    const title = template.title
    worksheet.mergeCells(currentRow, 1, currentRow, totalCols)
    const titleCell = worksheet.getCell(currentRow, 1)
    titleCell.value = title
    titleCell.style = titleStyle
    worksheet.getRow(currentRow).height = template.rowHeights.title || 69
    currentRow++
    
    // 生成注释
    const note = template.note
    worksheet.mergeCells(currentRow, 1, currentRow, totalCols)
    const noteCell = worksheet.getCell(currentRow, 1)
    noteCell.value = note
    noteCell.style = noteStyle
    worksheet.getRow(currentRow).height = template.rowHeights.note || 24.5
    currentRow++
    
    // 生成表头
    const headerRow = currentRow
    worksheet.getCell(headerRow, 1).value = '序号'
    worksheet.getCell(headerRow, 1).style = headerStyle
    worksheet.getCell(headerRow, 2).value = '班级'
    worksheet.getCell(headerRow, 2).style = headerStyle
    worksheet.getCell(headerRow, 3).value = '姓名'
    worksheet.getCell(headerRow, 3).style = headerStyle
    worksheet.getCell(headerRow, 4).value = '奖项'
    worksheet.getCell(headerRow, 4).style = headerStyle
    worksheet.getCell(headerRow, 5).value = '加分数'
    worksheet.getCell(headerRow, 5).style = headerStyle
    worksheet.getRow(headerRow).height = template.rowHeights.header || 20
    currentRow++
    
    // 生成数据行（单列）
    sortedData.forEach((row, index) => {
      worksheet.getCell(currentRow, 1).value = index + 1
      worksheet.getCell(currentRow, 1).style = bodyStyle
      worksheet.getCell(currentRow, 2).value = row['行政班级'] || ''
      worksheet.getCell(currentRow, 2).style = bodyStyle
      worksheet.getCell(currentRow, 3).value = row['姓名'] || ''
      worksheet.getCell(currentRow, 3).style = bodyStyle
      worksheet.getCell(currentRow, 4).value = row['奖项'] || ''
      worksheet.getCell(currentRow, 4).style = bodyStyle
      worksheet.getCell(currentRow, 5).value = row['加分分数'] || ''
      worksheet.getCell(currentRow, 5).style = bodyStyle
      worksheet.getRow(currentRow).height = template.rowHeights.body || 20
      currentRow++
    })
  } else {
    // 活动/讲座类型：双列布局（A、B、C列和D、E、F列）
    // 设置列宽
    worksheet.getColumn(1).width = 4.75  // A列：序号
    worksheet.getColumn(2).width = 27    // B列：行政班级
    worksheet.getColumn(3).width = 8.25  // C列：姓名
    worksheet.getColumn(4).width = 4.75  // D列：序号
    worksheet.getColumn(5).width = 27    // E列：行政班级
    worksheet.getColumn(6).width = 8.25 // F列：姓名
    
    const totalCols = 6
    
    // 生成标题
    const title = template.title
    worksheet.mergeCells(currentRow, 1, currentRow, totalCols)
    const titleCell = worksheet.getCell(currentRow, 1)
    titleCell.value = title
    titleCell.style = titleStyle
    worksheet.getRow(currentRow).height = template.rowHeights.title || 69
    currentRow++
    
    // 生成注释
    const note = template.note
    worksheet.mergeCells(currentRow, 1, currentRow, totalCols)
    const noteCell = worksheet.getCell(currentRow, 1)
    noteCell.value = note
    noteCell.style = noteStyle
    worksheet.getRow(currentRow).height = template.rowHeights.note || 24.5
    currentRow++
    
    // 生成表头（双列布局）
    const headerRow = currentRow
    // 第一列表头
    worksheet.getCell(headerRow, 1).value = '序号'
    worksheet.getCell(headerRow, 1).style = headerStyle
    worksheet.getCell(headerRow, 2).value = '班级'
    worksheet.getCell(headerRow, 2).style = headerStyle
    worksheet.getCell(headerRow, 3).value = '姓名'
    worksheet.getCell(headerRow, 3).style = headerStyle
    // 第二列表头
    worksheet.getCell(headerRow, 4).value = '序号'
    worksheet.getCell(headerRow, 4).style = headerStyle
    worksheet.getCell(headerRow, 5).value = '班级'
    worksheet.getCell(headerRow, 5).style = headerStyle
    worksheet.getCell(headerRow, 6).value = '姓名'
    worksheet.getCell(headerRow, 6).style = headerStyle
    worksheet.getRow(headerRow).height = template.rowHeights.header || 20
    currentRow++
    
    // 将数据分成两列显示，序号按"从上到下、从左到右"分配
    const totalRows = Math.ceil(sortedData.length / 2)
    
    for (let rowIndex = 0; rowIndex < totalRows; rowIndex++) {
      const dataRow = currentRow
      
      // 第一列数据（A、B、C列）
      const firstIndex = rowIndex * 2
      if (firstIndex < sortedData.length) {
        const firstData = sortedData[firstIndex]
        const firstSerial = firstIndex + 1
        
        worksheet.getCell(dataRow, 1).value = firstSerial
        worksheet.getCell(dataRow, 1).style = bodyStyle
        worksheet.getCell(dataRow, 2).value = firstData['行政班级'] || ''
        worksheet.getCell(dataRow, 2).style = bodyStyle
        worksheet.getCell(dataRow, 3).value = firstData['姓名'] || ''
        worksheet.getCell(dataRow, 3).style = bodyStyle
      }
      
      // 第二列数据（D、E、F列）
      const secondIndex = rowIndex * 2 + 1
      if (secondIndex < sortedData.length) {
        const secondData = sortedData[secondIndex]
        const secondSerial = secondIndex + 1
        
        worksheet.getCell(dataRow, 4).value = secondSerial
        worksheet.getCell(dataRow, 4).style = bodyStyle
        worksheet.getCell(dataRow, 5).value = secondData['行政班级'] || ''
        worksheet.getCell(dataRow, 5).style = bodyStyle
        worksheet.getCell(dataRow, 6).value = secondData['姓名'] || ''
        worksheet.getCell(dataRow, 6).style = bodyStyle
      }
      
      worksheet.getRow(dataRow).height = template.rowHeights.body || 20
      currentRow++
    }
  }
  
  // 导出文件
  const buffer = await workbook.xlsx.writeBuffer()
  const blob = new Blob([buffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' })
  const url = window.URL.createObjectURL(blob)
  const link = document.createElement('a')
  link.href = url
  link.download = filename
  document.body.appendChild(link)
  link.click()
  document.body.removeChild(link)
  window.URL.revokeObjectURL(url)
}

/**
 * 导出黑名单
 * @param {Array} data - 未考勤数据
 * @param {String} filename - 文件名
 */
export function exportBlacklist(data, filename) {
  const workbook = XLSX.utils.book_new()
  const headers = ['班级', '姓名', '学号', '未考勤原因']
  const rows = data.map(row => [
    row['行政班级'] || '',
    row['姓名'] || '',
    row['学号'] || '',
    row['未考勤原因'] || ''
  ])
  
  const worksheet = XLSX.utils.aoa_to_sheet([headers, ...rows])
  XLSX.utils.book_append_sheet(workbook, worksheet, '黑名单')
  XLSX.writeFile(workbook, filename)
}

