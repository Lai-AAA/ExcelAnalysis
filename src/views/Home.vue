<template>
  <div class="home-container">
    <el-card class="common-card">
      <template #header>
        <div class="card-header">
          <span>文件上传</span>
        </div>
      </template>
      
      <el-upload
        ref="uploadRef"
        class="upload-demo"
        drag
        :auto-upload="false"
        :limit="3"
        :on-exceed="handleExceed"
        :on-change="handleFileChange"
        :file-list="fileList"
        accept=".xlsx"
        :before-upload="beforeUpload"
      >
        <el-icon class="el-icon--upload"><upload-filled /></el-icon>
        <div class="el-upload__text">
          将文件拖到此处，或<em>点击上传</em>
        </div>
        <template #tip>
          <div class="el-upload__tip">
            支持.xlsx格式，单个文件不超过10MB，最多同时上传3个文件
          </div>
        </template>
      </el-upload>
      
      <div class="upload-actions" v-if="fileList.length > 0">
        <el-button type="primary" @click="handleUpload" :loading="uploading">
          开始解析
        </el-button>
        <el-button @click="handleClear">清空文件</el-button>
      </div>
    </el-card>
    
    <el-card class="common-card">
      <template #header>
        <div class="card-header">
          <span>操作指南</span>
        </div>
      </template>
      
      <el-steps :active="0" finish-status="success" align-center>
        <el-step title="上传文件" description="上传包含'数据源'工作表的Excel文件" />
        <el-step title="数据解析" description="自动解析数据并验证考勤" />
        <el-step title="数据筛选" description="按条件筛选需要导出的数据" />
        <el-step title="导出名单" description="按模板格式导出加分名单" />
      </el-steps>
    </el-card>
    
    <el-card class="common-card">
      <template #header>
        <div class="card-header">
          <span>Python脚本下载</span>
        </div>
      </template>
      
      <div class="script-info">
        <p>针对大数据量场景（≥3000行），可使用Python脚本提升处理效率：</p>
        <ul>
          <li><strong>data_parser.py</strong> - 数据解析与考勤验证</li>
          <li><strong>data_filter.py</strong> - 多维度数据筛选</li>
          <li><strong>list_generator.py</strong> - 加分名单生成</li>
        </ul>
        <el-button type="primary" @click="downloadScripts">
          <el-icon><Download /></el-icon>
          下载Python脚本包
        </el-button>
      </div>
    </el-card>
  </div>
</template>

<script setup>
import { ref } from 'vue'
import { useRouter } from 'vue-router'
import { ElMessage, ElMessageBox } from 'element-plus'
import { UploadFilled, Download } from '@element-plus/icons-vue'
import { parseExcelFile } from '@/utils/excel'
import { saveData } from '@/utils/storage'

const router = useRouter()
const uploadRef = ref(null)
const fileList = ref([])
const uploading = ref(false)

const beforeUpload = (file) => {
  const isXLSX = file.name.endsWith('.xlsx')
  const isLt10M = file.size / 1024 / 1024 < 10
  
  if (!isXLSX) {
    ElMessage.error('请上传标准Excel文件（.xlsx）')
    return false
  }
  if (!isLt10M) {
    ElMessage.error('文件大小不能超过10MB')
    return false
  }
  return false // 阻止自动上传
}

const handleExceed = () => {
  ElMessage.warning('最多只能上传3个文件')
}

const handleFileChange = (file, files) => {
  fileList.value = files
}

const handleClear = () => {
  fileList.value = []
  uploadRef.value?.clearFiles()
}

const handleUpload = async () => {
  if (fileList.value.length === 0) {
    ElMessage.warning('请先选择文件')
    return
  }
  
  uploading.value = true
  
  try {
    const allData = []
    const allStats = {
      total: 0,
      valid: 0,
      invalid: 0,
      classes: new Set(),
      activityTypes: new Set()
    }
    
    // 解析所有文件
    for (const file of fileList.value) {
      const result = await parseExcelFile(file.raw)
      allData.push(...result.processedData)
      
      allStats.total += result.stats.total
      allStats.valid += result.stats.valid
      allStats.invalid += result.stats.invalid
      result.processedData.forEach(row => {
        allStats.classes.add(row['行政班级'])
        allStats.activityTypes.add(row['活动类型'])
      })
    }
    
    // 保存数据
    saveData({
      rawData: allData,
      stats: {
        ...allStats,
        classes: allStats.classes.size,
        activityTypes: allStats.activityTypes.size
      }
    })
    
    ElMessage.success(`成功解析${allStats.valid}行有效数据`)
    
    // 跳转到解析页
    router.push('/parse')
  } catch (error) {
    ElMessage.error(error.message || '文件解析失败')
  } finally {
    uploading.value = false
  }
}

const downloadScripts = () => {
  ElMessage.info('Python脚本位于项目根目录的scripts文件夹中')
}
</script>

<style lang="scss" scoped>
.home-container {
  max-width: 1200px;
  margin: 0 auto;
}

.upload-demo {
  :deep(.el-upload-dragger) {
    width: 100%;
  }
}

.upload-actions {
  margin-top: 20px;
  text-align: center;
}

.script-info {
  ul {
    margin: 15px 0;
    padding-left: 20px;
    
    li {
      margin: 8px 0;
    }
  }
}
</style>

