<script setup lang="ts">
import { RouterView } from 'vue-router'

</script>

<template>
  <div class="min-h-screen bg-gray-50 text-gray-900">
    <header class="bg-white shadow px-4 py-3">
      <div class="max-w-4xl mx-auto flex justify-between items-center">
        <h1 class="text-xl font-semibold">議事録作成ツール</h1>
        <nav class="space-x-4 text-sm">
          <RouterLink to="/" class="hover:underline">議事録作成</RouterLink>
          <RouterLink to="/grouped" class="hover:underline">議事録作成(議題グループ)</RouterLink>
          <RouterLink to="/about" class="hover:underline">このツールについて</RouterLink>
          <RouterLink to="/help" class="hover:underline">使い方</RouterLink>
          <RouterLink to="/terms" class="hover:underline">利用規約</RouterLink>
        </nav>
      </div>
    </header>

    <main class="max-w-4xl mx-auto px-4 py-6">
      <RouterView />
    </main>

    <footer class="text-center text-xs text-gray-500 py-4 border-t">
      © 2025 Takahiro Yamashita
    </footer>
  </div>
</template>

<style scoped>
</style>
