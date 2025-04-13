import { createRouter, createWebHistory } from 'vue-router'
import HomeView from '../views/HomeView.vue'
import HomeGroupedView from '../views/HomeGroupedView.vue'
import AboutView from '../views/AboutView.vue'
import TermsView from '../views/TermsView.vue'
import HelpView from '../views/HelpView.vue'

const routes = [
  { path: '/', name: 'Home', component: HomeView },
  { path: '/grouped', name: 'Grouped', component: HomeGroupedView },
  { path: '/about', name: 'About', component: AboutView },
  { path: '/terms', name: 'Terms', component: TermsView },
  { path: '/help', name: 'Help', component: HelpView },
]

const router = createRouter({
  history: createWebHistory(),
  routes,
})

export default router
