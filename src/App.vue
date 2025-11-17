<template>
  <el-container class="app-container">
    <el-header class="app-header">
      <h1>Vue加分名单生成系统</h1>
    </el-header>
    <el-main class="app-main">
      <router-view />
    </el-main>
  </el-container>
</template>

<script setup>
</script>

<style lang="scss" scoped>
.app-container {
  min-height: 100vh;
}

.app-header {
  background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
  color: white;
  display: flex;
  align-items: center;
  justify-content: center;
  box-shadow: 0 2px 8px rgba(0, 0, 0, 0.1);

  h1 {
    margin: 0;
    font-size: 24px;
    font-weight: 500;
  }
}

.app-main {
  padding: 20px;
  background-color: #f5f7fa;
}
</style>

