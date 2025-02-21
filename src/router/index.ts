import { createRouter, createWebHashHistory } from 'vue-router'
// 自动导入路由
import { routes } from 'vue-router/auto-routes'

const router = createRouter({
  history: createWebHashHistory(),
  routes,
})

export default router
