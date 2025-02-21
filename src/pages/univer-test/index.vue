<script lang="ts" setup>
import type { FUniver, Univer } from '@univerjs/presets'
import type { FWorkbook, FWorksheet } from '@univerjs/presets/preset-sheets-core'
import { createUniver, defaultTheme, LocaleType, merge } from '@univerjs/presets'
import { UniverSheetsCorePreset } from '@univerjs/presets/preset-sheets-core'
import UniverPresetSheetsCoreZhCN from '@univerjs/presets/preset-sheets-core/locales/zh-CN'
import axios from 'axios'

import { ElMessage } from 'element-plus'
import { onMounted, reactive, ref } from 'vue'
import { useRouter } from 'vue-router'
import '@univerjs/presets/lib/styles/preset-sheets-core.css'

const sheetContainerRef = ref()

const router = useRouter()

const loading = reactive({
  chatLoading: false,
})

const formValues = reactive({
  target: '',
  targetValue: '',
  message: '填充A1-C8值为Lucky，每隔一个空出来',
})

const sheetApplication = {
  univerAPI: {} as FUniver,
  univer: {} as Univer,
  wordbook: {} as FWorkbook,
  wordSheet: {} as FWorksheet,
  init() {
    // 实例化 univer
    const { univer, univerAPI } = createUniver({
      locale: LocaleType.ZH_CN,
      locales: {
        [LocaleType.ZH_CN]: merge(
          {},
          UniverPresetSheetsCoreZhCN,
        ),
      },
      theme: defaultTheme,
      presets: [
        UniverSheetsCorePreset({
          container: sheetContainerRef.value,
        }),
      ],
    })
    sheetApplication.univerAPI = univerAPI
    sheetApplication.univer = univer
    // 创建表格
    sheetApplication.wordbook = sheetApplication.univerAPI.createWorkbook({
      name: 'sheet test',
    })
    sheetApplication.wordSheet = sheetApplication.wordbook.getActiveSheet()
  },
  checkRange() {
    const cellRegex = /^[A-Z]{1,3}\d{1,7}$/i
    return cellRegex.test(formValues.target)
  },
  setTargetValue() {
    const { target, targetValue } = formValues
    if (!sheetApplication.checkRange()) {
      ElMessage.error('目标单元格输入不正确，请检查')
      return
    }
    sheetApplication.wordSheet.getRange(target).setValue(targetValue)
  },
}

const aiApplication = {
  async chat() {
    loading.chatLoading = true
    const res = await axios.post('/api/chat', {
      message: formValues.message,
    })
    loading.chatLoading = false
    const { code, message } = res.data
    if (code !== 0) {
      ElMessage.error(code)
    }
    const instructions = JSON.parse(message.content)
    for (const item of instructions) {
      sheetApplication.wordSheet.getRange(item[0]).setValue(item[1])
    }
  },
}

function changPage() {
  router.push('/lucky-sheet-test')
}

onMounted(() => {
  sheetApplication.init()
})
</script>

<template>
  <div class="page-container">
    <div class="operate-container">
      <el-form :inline="true" :model="formValues" class="demo-form-inline">
        <el-form-item label="单元格">
          <el-input v-model="formValues.target" class="common-input" width="220" placeholder="目标单元格" clearable />
        </el-form-item>
        <el-form-item label="单元格值">
          <el-input v-model="formValues.targetValue" class="common-input" placeholder="请输入单元格值" clearable />
        </el-form-item>
        <el-form-item>
          <el-button type="primary" @click="sheetApplication.setTargetValue()">
            确认
          </el-button>
          <el-button type="warning" @click="changPage()">
            切换 LuckySheet
          </el-button>
        </el-form-item>
        <el-form-item style="margin-left: 50px;" label="DeepSeek">
          <el-input v-model="formValues.message" type="textarea" style="width: 320px;" :rows="2" placeholder="" clearable />
        </el-form-item>
        <el-form-item>
          <el-button :loading="loading.chatLoading" type="warning" @click="aiApplication.chat()">
            执行
          </el-button>
        </el-form-item>
      </el-form>
    </div>
    <div ref="sheetContainerRef" class="sheet-container" />
  </div>
</template>

<style lang="scss" scoped>
.page-container {
  display: flex;
  align-items: center;
  justify-content: center;
  flex-direction: column;
  height: 100vh;
  width: 100vw;

  .operate-container {
    display: flex;
    align-items: center;
    justify-content: flex-start;
    width: 100%;
    padding: 20px;
  }

  .sheet-container {
    width: 100%;
    height: 200px;
    flex: 1;
  }
}

.common-input {
  width: 220px;
}
</style>
