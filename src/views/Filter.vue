<template>
  <div class="filter-container">
    <el-card class="common-card">
      <template #header>
        <div class="card-header">
          <span>数据筛选</span>
        </div>
      </template>
      
      <el-form :model="filters" label-width="120px">
        <el-row :gutter="20">
          <el-col :span="12">
            <el-form-item label="学年学期">
              <el-select v-model="filters.semester" placeholder="请选择学年学期" clearable style="width: 100%">
                <el-option
                  v-for="semester in semesters"
                  :key="semester"
                  :label="semester"
                  :value="semester"
                />
              </el-select>
            </el-form-item>
          </el-col>
          
          <el-col :span="12">
            <el-form-item label="活动类型">
              <el-select
                v-model="filters.activityTypes"
                placeholder="请选择活动类型"
                multiple
                clearable
                style="width: 100%"
              >
                <el-option
                  v-for="type in activityTypes"
                  :key="type"
                  :label="type"
                  :value="type"
                />
              </el-select>
            </el-form-item>
          </el-col>
        </el-row>
        
        <el-row :gutter="20">
          <el-col :span="12">
            <el-form-item label="活动名称">
              <el-input
                v-model="filters.activityName"
                placeholder="支持模糊搜索"
                clearable
              />
            </el-form-item>
          </el-col>
          
          <el-col :span="12">
            <el-form-item label="加分类型">
              <el-select v-model="filters.scoreType" placeholder="请选择加分类型" clearable style="width: 100%">
                <el-option label="学业分" value="学业分" />
                <el-option label="品德分" value="品德分" />
              </el-select>
            </el-form-item>
          </el-col>
        </el-row>
        
        <el-row :gutter="20">
          <el-col :span="12">
            <el-form-item label="加分分数范围">
              <el-input-number
                v-model="filters.minScore"
                :min="0"
                :max="10"
                placeholder="最小值"
                style="width: 48%"
              />
              <span style="margin: 0 2%">-</span>
              <el-input-number
                v-model="filters.maxScore"
                :min="0"
                :max="10"
                placeholder="最大值"
                style="width: 48%"
              />
            </el-form-item>
          </el-col>
          
          <el-col :span="12">
            <el-form-item label="行政班级">
              <el-select
                v-model="filters.classes"
                placeholder="请选择班级（可多选）"
                multiple
                filterable
                clearable
                style="width: 100%"
              >
                <el-option
                  v-for="cls in classes"
                  :key="cls"
                  :label="cls"
                  :value="cls"
                />
              </el-select>
            </el-form-item>
          </el-col>
        </el-row>
        
        <el-row :gutter="20">
          <el-col :span="12">
            <el-form-item label="其他选项">
              <el-checkbox v-model="filters.includeOtherCollege">包含其他学院人员</el-checkbox>
              <el-checkbox v-model="filters.sortByClass" :true-label="true" :false-label="false">按班级排序</el-checkbox>
            </el-form-item>
          </el-col>
        </el-row>
        
        <el-form-item>
          <el-button type="primary" @click="handleFilter">应用筛选</el-button>
          <el-button @click="handleReset">重置筛选</el-button>
        </el-form-item>
      </el-form>
    </el-card>
    
    <el-card class="common-card" v-if="filteredData.length > 0">
      <template #header>
        <div class="card-header">
          <span>筛选结果（共{{ filteredData.length }}条）</span>
        </div>
      </template>
      
      <el-table :data="filteredData.slice(0, 20)" border style="width: 100%">
        <el-table-column prop="行政班级" label="班级" width="150" />
        <el-table-column prop="姓名" label="姓名" width="120" />
        <el-table-column prop="活动名称" label="活动名称" width="200" />
        <el-table-column prop="活动类型" label="活动类型" width="100" />
        <el-table-column prop="加分类型" label="加分类型" width="100" />
        <el-table-column prop="加分分数" label="加分分数" width="100" />
        <el-table-column prop="奖项" label="奖项" width="100" />
      </el-table>
      
      <div class="table-footer" v-if="filteredData.length > 20">
        <el-text>共{{ filteredData.length }}条记录，仅显示前20条</el-text>
      </div>
    </el-card>
    
    <div class="action-buttons" v-if="filteredData.length > 0">
      <el-button type="primary" size="large" @click="goToExport">
        进入导出页面
      </el-button>
    </div>
  </div>
</template>

<script setup>
import { ref, onMounted } from 'vue'
import { useRouter } from 'vue-router'
import { ElMessage } from 'element-plus'
import { loadData, saveData } from '@/utils/storage'
import { filterData, getUniqueSemesters, getUniqueActivityTypes, getUniqueClasses } from '@/utils/filter'

const router = useRouter()
const sourceData = ref([])
const filteredData = ref([])
const semesters = ref([])
const activityTypes = ref([])
const classes = ref([])

const filters = ref({
  semester: '',
  activityTypes: [],
  activityName: '',
  scoreType: '',
  minScore: null,
  maxScore: null,
  classes: [],
  includeOtherCollege: true,
  sortByClass: true
})

onMounted(() => {
  const data = loadData()
  if (!data || !data.attendedData) {
    ElMessage.warning('请先完成考勤验证')
    router.push('/parse')
    return
  }
  
  sourceData.value = data.attendedData
  
  // 初始化筛选选项
  semesters.value = getUniqueSemesters(sourceData.value)
  activityTypes.value = getUniqueActivityTypes(sourceData.value)
  classes.value = getUniqueClasses(sourceData.value)
  
  // 默认显示所有数据
  filteredData.value = sourceData.value
})

const handleFilter = () => {
  filteredData.value = filterData(sourceData.value, filters.value)
  
  // 保存筛选结果
  saveData({
    ...loadData(),
    filteredData: filteredData.value,
    filters: filters.value
  })
  
  ElMessage.success(`筛选完成，共找到${filteredData.value.length}条数据`)
}

const handleReset = () => {
  filters.value = {
    semester: '',
    activityTypes: [],
    activityName: '',
    scoreType: '',
    minScore: null,
    maxScore: null,
    classes: [],
    includeOtherCollege: true,
    sortByClass: true
  }
  filteredData.value = sourceData.value
  ElMessage.info('筛选条件已重置')
}

const goToExport = () => {
  router.push('/export')
}
</script>

<style lang="scss" scoped>
.filter-container {
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
</style>

