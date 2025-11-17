import { createRouter, createWebHistory } from 'vue-router'
import Home from '../views/Home.vue'
import Parse from '../views/Parse.vue'
import Filter from '../views/Filter.vue'
import Export from '../views/Export.vue'

const routes = [
  {
    path: '/',
    name: 'Home',
    component: Home
  },
  {
    path: '/parse',
    name: 'Parse',
    component: Parse
  },
  {
    path: '/filter',
    name: 'Filter',
    component: Filter
  },
  {
    path: '/export',
    name: 'Export',
    component: Export
  }
]

const router = createRouter({
  history: createWebHistory(),
  routes
})

export default router

