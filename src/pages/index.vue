<script setup lang="ts">
import { useRouter } from 'vue-router'

const router = useRouter()

const goPage = (page: string) => router.replace(page)
</script>

<template>
  <div class="describe">
    <h1>不是很精美的Demo，使用univer 又顺手写了个LuckySheet版本</h1>
    <h2>心血来潮接入了DeepSeek,只不过只能做简单的事情（时间问题只能写不能读）</h2>
    <h3>
      univer是新版本，TypeScript支持较好。但是文档不够完善，接口不够全，比如没有修改单元格格式的接口（也可能是没找到）；目前不推荐升级
    </h3>
    <h3>
      LuckySheet版本功能较全；只是没有TypeScript支持，不在维护了，好在也足够稳定
    </h3>
    <div class="operate-bar">
      <el-button type="primary" @click=" goPage('/lucky-sheet-test')">
        LuckySheet
      </el-button>
      <el-button type="warning" @click="goPage('/univer-test')">
        Univer
      </el-button>
    </div>
  </div>
</template>

<style lang="scss" scoped>
.describe {
  width: 100%;
  height: 100vh;
  display: flex;
  align-items: center;
  justify-content: center;
  flex-direction: column;
  .operate-bar {
    padding-top: 50px;
  }
}
</style>
