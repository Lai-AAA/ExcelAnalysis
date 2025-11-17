<template>
  <div class="parse-container">
    <el-card class="common-card">
      <template #header>
        <div class="card-header">
          <span>数据统计</span>
        </div>
      </template>
      
      <el-row :gutter="20" v-if="stats">
        <el-col :span="6">
          <div class="stat-card">
            <div class="stat-label">总数据</div>
            <div class="stat-value">{{ stats.total }}</div>
          </div>
        </el-col>
        <el-col :span="6">
          <div class="stat-card" style="background: linear-gradient(135deg, #11998e 0%, #38ef7d 100%);">
            <div class="stat-label">有效数据</div>
            <div class="stat-value">{{ stats.valid }}</div>
          </div>
        </el-col>
        <el-col :span="6">
          <div class="stat-card" style="background: linear-gradient(135deg, #f093fb 0%, #f5576c 100%);">
            <div class="stat-label">涉及班级</div>
            <div class="stat-value">{{ stats.classes }}</div>
          </div>
        </el-col>
        <el-col :span="6">
          <div class="stat-card" style="background: linear-gradient(135deg, #4facfe 0%, #00f2fe 100%);">
            <div class="stat-label">活动类型</div>
            <div class="stat-value">{{ stats.activityTypes }}</div>
          </div>
        </el-col>
      </el-row>
    </el-card>
    
    <el-card class="common-card">
      <template #header>
        <div class="card-header">
          <span>考勤验证</span>
        </div>
      </template>
      
      <el-form :model="attendanceForm" label-width="120px">
        <el-form-item label="考勤关键词">
          <el-input
            v-model="attendanceForm.keywords"
            placeholder="用逗号分隔，如：已考勤,到场,参与"
          />
        </el-form-item>
        <el-form-item>
          <el-button type="primary" @click="handleValidateAttendance">验证考勤</el-button>
        </el-form-item>
      </el-form>
      
      <el-row :gutter="20" v-if="attendanceStats">
        <el-col :span="12">
          <div class="attendance-stat attended">
            <div class="stat-label">已考勤（可加分）</div>
            <div class="stat-value">{{ attendanceStats.attended }}</div>
          </div>
        </el-col>
        <el-col :span="12">
          <div class="attendance-stat unattended">
            <div class="stat-label">未考勤（黑名单）</div>
            <div class="stat-value">{{ attendanceStats.unattended }}</div>
          </div>
        </el-col>
      </el-row>
    </el-card>
    
    <el-card class="common-card" v-if="blacklist.length > 0">
      <template #header>
        <div class="card-header">
          <span>黑名单预览</span>
          <el-button type="danger" size="small" @click="handleExportBlacklist">
            导出黑名单
          </el-button>
        </div>
      </template>
      
      <el-table :data="blacklist.slice(0, 10)" border style="width: 100%">
        <el-table-column prop="行政班级" label="班级" width="150" />
        <el-table-column prop="姓名" label="姓名" width="120" />
        <el-table-column prop="学号" label="学号" width="150" />
        <el-table-column prop="未考勤原因" label="未考勤原因" />
      </el-table>
      
      <div class="table-footer" v-if="blacklist.length > 10">
        <el-text>共{{ blacklist.length }}条记录，仅显示前10条</el-text>
      </div>
    </el-card>
    
    <el-card class="common-card" v-if="attendedData.length > 0">
      <template #header>
        <div class="card-header">
          <span>已考勤数据预览</span>
        </div>
      </template>
      
      <el-table :data="attendedData.slice(0, 10)" border style="width: 100%">
        <el-table-column prop="行政班级" label="班级" width="150" />
        <el-table-column prop="姓名" label="姓名" width="120" />
        <el-table-column prop="活动名称" label="活动名称" width="200" />
        <el-table-column prop="活动类型" label="活动类型" width="100" />
        <el-table-column prop="加分类型" label="加分类型" width="100" />
        <el-table-column prop="加分分数" label="加分分数" width="100" />
      </el-table>
      
      <div class="table-footer" v-if="attendedData.length > 10">
        <el-text>共{{ attendedData.length }}条记录，仅显示前10条</el-text>
      </div>
    </el-card>
    
    <div class="action-buttons" v-if="attendedData.length > 0">
      <el-button type="primary" size="large" @click="goToFilter">
        进入数据筛选
      </el-button>
    </div>
  </div>
</template>

<script setup>
import { ref, onMounted } from 'vue'
import { useRouter } from 'vue-router'
import { ElMessage } from 'element-plus'
import { loadData, saveData } from '@/utils/storage'
import { validateAttendance } from '@/utils/excel'
import { exportBlacklist } from '@/utils/excel'

const router = useRouter()
const stats = ref(null)
const rawData = ref([])
const attendanceForm = ref({
  keywords: '已考勤,到场,参与'
})
const attendedData = ref([])
const blacklist = ref([])
const attendanceStats = ref(null)

onMounted(() => {
  const data = loadData()
  if (!data || !data.rawData) {
    ElMessage.warning('请先上传并解析文件')
    router.push('/')
    return
  }
  
  rawData.value = data.rawData
  stats.value = data.stats
  
  // 自动执行一次考勤验证
  handleValidateAttendance()
})

const handleValidateAttendance = () => {
  const keywords = attendanceForm.value.keywords
    .split(',')
    .map(k => k.trim())
    .filter(k => k)
  
  if (keywords.length === 0) {
    ElMessage.warning('请输入考勤关键词')
    return
  }
  
  const result = validateAttendance(rawData.value, { keywords })
  attendedData.value = result.attended
  blacklist.value = result.unattended
  
  attendanceStats.value = {
    attended: result.attended.length,
    unattended: result.unattended.length
  }
  
  // 保存验证后的数据
  saveData({
    rawData: rawData.value,
    stats: stats.value,
    attendedData: result.attended,
    blacklist: result.unattended
  })
  
  ElMessage.success('考勤验证完成')
}

const handleExportBlacklist = () => {
  if (blacklist.value.length === 0) {
    ElMessage.warning('没有黑名单数据可导出')
    return
  }
  
  // 获取学年学期（从第一条数据获取）
  const semester = rawData.value[0]?.['学年学期'] || '2024-2025学年第X学期'
  const filename = `${semester}未考勤人员黑名单.xlsx`
  
  exportBlacklist(blacklist.value, filename)
  ElMessage.success('黑名单导出成功')
}

const goToFilter = () => {
  router.push('/filter')
}
</script>

<style lang="scss" scoped>
.parse-container {
  max-width: 1400px;
  margin: 0 auto;
}

.attendance-stat {
  padding: 20px;
  border-radius: 8px;
  text-align: center;
  color: white;
  
  &.attended {
    background: linear-gradient(135deg, #11998e 0%, #38ef7d 100%);
  }
  
  &.unattended {
    background: linear-gradient(135deg, #f093fb 0%, #f5576c 100%);
  }
  
  .stat-label {
    font-size: 14px;
    opacity: 0.9;
    margin-bottom: 10px;
  }
  
  .stat-value {
    font-size: 32px;
    font-weight: bold;
  }
}

.table-footer {
  margin-top: 15px;
  text-align: center;
}

.action-buttons {
  text-align: center;
  margin-top: 30px;
}
</style>

