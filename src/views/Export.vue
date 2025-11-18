<template>
  <div class="export-container">
    <el-card class="common-card">
      <template #header>
        <div class="card-header">
          <span>模板选择与预览</span>
        </div>
      </template>
      
      <el-form :model="exportForm" label-width="120px">
        <el-form-item label="活动类型">
          <el-radio-group v-model="exportForm.activityType">
            <el-radio label="活动">活动/讲座类</el-radio>
            <el-radio label="比赛">比赛类</el-radio>
          </el-radio-group>
        </el-form-item>
        
        <el-form-item label="活动名称" v-if="exportForm.activityType">
          <el-input
            v-model="exportForm.activityName"
            placeholder="请输入活动名称"
          />
        </el-form-item>
        
        <el-form-item label="学年学期">
          <el-select v-model="exportForm.semester" placeholder="请选择学年学期" style="width: 100%">
            <el-option
              v-for="semester in semesters"
              :key="semester"
              :label="semester"
              :value="semester"
            />
          </el-select>
        </el-form-item>
        
        <el-form-item label="加分类型">
          <el-select v-model="exportForm.scoreType" placeholder="请选择加分类型" style="width: 100%">
            <el-option label="学业分" value="学业分" />
            <el-option label="品德分" value="品德分" />
          </el-select>
        </el-form-item>
        
        <el-form-item label="添加注释" v-if="exportForm.activityType">
          <el-input
            v-model="exportForm.customNote"
            type="textarea"
            :rows="3"
            placeholder="请输入注释内容（可选）"
            style="width: 100%"
          />
          <div class="note-template-hint">
            注：如对以上有异议请联系工学院团委/学生会XX（部门）干部XXX（联系方式）
          </div>
        </el-form-item>
      </el-form>
    </el-card>
    
    <el-card class="common-card" v-if="previewData.length > 0">
      <template #header>
        <div class="card-header">
          <span>导出预览（共{{ previewData.length }}条）</span>
        </div>
      </template>
      
      <el-table :data="previewData.slice(0, 10)" border style="width: 100%">
        <el-table-column prop="序号" label="序号" width="80" />
        <el-table-column prop="行政班级" label="班级" width="150" />
        <el-table-column prop="姓名" label="姓名" width="120" />
        <el-table-column prop="奖项" label="奖项" width="100" v-if="exportForm.activityType === '比赛'" />
        <el-table-column prop="加分分数" label="加分分数" width="100" v-if="exportForm.activityType === '比赛'" />
      </el-table>
      
      <div class="table-footer" v-if="previewData.length > 10">
        <el-text>共{{ previewData.length }}条记录，仅显示前10条</el-text>
      </div>
    </el-card>
    
    <div class="action-buttons">
      <el-button type="primary" size="large" @click="handleExport" :disabled="!canExport">
        导出加分名单
      </el-button>
      <el-button size="large" @click="goBack">返回筛选</el-button>
    </div>
  </div>
</template>

<script setup>
import { ref, computed, onMounted, watch } from 'vue'
import { useRouter } from 'vue-router'
import { ElMessage } from 'element-plus'
import { loadData, saveExportHistory } from '@/utils/storage'
import { exportExcel } from '@/utils/excel'
import { getTemplateByActivityType } from '@/utils/templates'
import { getUniqueSemesters } from '@/utils/filter'

const router = useRouter()
const filteredData = ref([])
const semesters = ref([])
const sourceData = ref(null)

const exportForm = ref({
  activityType: '活动',
  activityName: '',
  semester: '',
  scoreType: '',
  customNote: ''
})

const previewData = ref([])

const canExport = computed(() => {
  return (
    exportForm.value.activityType &&
    exportForm.value.activityName &&
    exportForm.value.semester &&
    exportForm.value.scoreType &&
    previewData.value.length > 0
  )
})

const updatePreview = () => {
  if (!exportForm.value.activityType || !filteredData.value.length) {
    previewData.value = []
    return
  }
  
  // 根据活动类型筛选
  let data = filteredData.value
  if (exportForm.value.activityName) {
    data = data.filter(row => 
      row['活动名称'] === exportForm.value.activityName
    )
  }
  
  // 添加序号
  previewData.value = data.map((row, index) => ({
    ...row,
    序号: index + 1
  }))
}

watch(() => exportForm.value.activityType, updatePreview)
watch(() => exportForm.value.activityName, updatePreview)

onMounted(() => {
  const data = loadData()
  if (!data || !data.filteredData) {
    ElMessage.warning('请先完成数据筛选')
    router.push('/filter')
    return
  }
  
  filteredData.value = data.filteredData
  sourceData.value = data
  
  // 初始化选项
  semesters.value = getUniqueSemesters(filteredData.value)
  if (semesters.value.length > 0) {
    exportForm.value.semester = semesters.value[0]
  }
  
  // 获取活动类型和名称（从第一条数据）
  if (filteredData.value.length > 0) {
    const firstRow = filteredData.value[0]
    exportForm.value.activityType = firstRow['活动类型'] || '活动'
    exportForm.value.activityName = firstRow['活动名称'] || ''
    exportForm.value.scoreType = firstRow['加分类型'] || '学业分'
  }
  
  updatePreview()
})

const handleExport = async () => {
  if (!canExport.value) {
    ElMessage.warning('请完善导出信息')
    return
  }
  
  try {
    const template = getTemplateByActivityType(exportForm.value.activityType)
    const semester = exportForm.value.semester
    const activityName = exportForm.value.activityName
    
    // 生成标题和注释
    const title = template.generateTitle(semester, activityName)
    const note = exportForm.value.activityType === '比赛'
      ? template.generateNote(exportForm.value.scoreType)
      : template.generateNote(exportForm.value.scoreType, exportForm.value.score)
    
    // 准备导出数据
    const exportData = previewData.value.map(row => {
      const item = { ...row }
      // 移除序号字段（会在导出时重新生成）
      delete item.序号
      return item
    })
    
    // 生成文件名
    const filename = `${semester}工学院${activityName}加分名单.xlsx`
    
    // 导出
    await exportExcel(exportData, {
      ...template,
      title,
      note,
      customNote: exportForm.value.customNote || '注：如对以上有异议请联系工学院团委/学生会XX（部门）干部XXX（联系方式）'
    }, filename, exportForm.value.activityType)
    
    // 保存导出历史
    saveExportHistory({
      filename,
      count: exportData.length,
      activityType: exportForm.value.activityType,
      activityName: exportForm.value.activityName
    })
    
    ElMessage.success('导出成功！')
  } catch (error) {
    ElMessage.error(`导出失败：${error.message}`)
  }
}

const goBack = () => {
  router.push('/filter')
}
</script>

<style lang="scss" scoped>
.export-container {
  max-width: 1400px;
  margin: 0 auto;
}

.table-footer {
  margin-top: 15px;
  text-align: center;
}

.action-buttons {
  text-align: center;
  margin-top: 30px;
}

.note-template-hint {
  margin-top: 8px;
  font-size: 10px;
  color: #000;
  line-height: 1.5;
}
</style>

