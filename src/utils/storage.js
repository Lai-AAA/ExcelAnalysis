const STORAGE_KEY = 'score_list_generator_data'
const HISTORY_KEY = 'score_list_generator_history'
const MAX_HISTORY = 3

/**
 * 保存数据到localStorage
 */
export function saveData(data) {
  try {
    localStorage.setItem(STORAGE_KEY, JSON.stringify(data))
  } catch (error) {
    console.error('保存数据失败:', error)
  }
}

/**
 * 从localStorage读取数据
 */
export function loadData() {
  try {
    const data = localStorage.getItem(STORAGE_KEY)
    return data ? JSON.parse(data) : null
  } catch (error) {
    console.error('读取数据失败:', error)
    return null
  }
}

/**
 * 清除数据
 */
export function clearData() {
  try {
    localStorage.removeItem(STORAGE_KEY)
  } catch (error) {
    console.error('清除数据失败:', error)
  }
}

/**
 * 保存导出历史
 */
export function saveExportHistory(record) {
  try {
    const history = getExportHistory()
    history.unshift({
      ...record,
      timestamp: new Date().toISOString()
    })
    // 只保留最近3条
    if (history.length > MAX_HISTORY) {
      history.pop()
    }
    localStorage.setItem(HISTORY_KEY, JSON.stringify(history))
  } catch (error) {
    console.error('保存导出历史失败:', error)
  }
}

/**
 * 获取导出历史
 */
export function getExportHistory() {
  try {
    const history = localStorage.getItem(HISTORY_KEY)
    return history ? JSON.parse(history) : []
  } catch (error) {
    console.error('读取导出历史失败:', error)
    return []
  }
}

