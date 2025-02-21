<template>
  <el-config-provider namespace="ep">
    <RouterView />
  </el-config-provider>
</template>

<style lang="scss" scoped>
@import url('./styles/index.scss');

#app {
  text-align: center;
  color: var(--ep-text-color-primary);
}

.main-container {
  height: calc(100vh - var(--ep-menu-item-height) - 4px);
}
</style>
