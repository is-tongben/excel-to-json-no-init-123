<script setup>
import { ref } from 'vue'

const count = ref(0)
</script>

<template>
  <div class="app">
    <h1>Vite + Vue</h1>
    <div class="card">
      <button type="button" @click="count++">count is {{ count }}</button>
      <p>
        编辑
        <code>src/App.vue</code>
        并保存以测试热更新。
      </p>
    </div>

    <!-- 图片网格展示 -->
    <div class="image-grid">
      <img src="/images/a.webp" alt="图片1" />
      <img src="/images/b.webp" alt="图片2" />
      <img src="/images/c.png" alt="图片3" />
      <img src="/images/d.png" alt="图片3" />
    </div>

    <p class="read-the-docs">
      点击 Vite 和 Vue 徽标以了解更多
    </p>
  </div>
</template>

<style scoped>
.app {
  font-family: Avenir, Helvetica, Arial, sans-serif;
  text-align: center;
  color: #2c3e50;
  margin-top: 60px;
}

.logo {
  height: 6em;
  padding: 1.5em;
  will-change: filter;
  transition: filter 300ms;
}
.logo:hover {
  filter: drop-shadow(0 0 2em #646cffaa);
}
.logo.vue:hover {
  filter: drop-shadow(0 0 2em #42b883aa);
}

.card {
  padding: 2em;
}

button {
  border-radius: 8px;
  border: 1px solid transparent;
  padding: 0.6em 1.2em;
  font-size: 1em;
  font-weight: 500;
  font-family: inherit;
  background-color: #1a1a1a;
  cursor: pointer;
  transition: border-color 0.25s;
}
button:hover {
  border-color: #646cff;
}
button:focus,
button:focus-visible {
  outline: 4px auto -webkit-focus-ring-color;
}

@media (prefers-color-scheme: light) {
  button {
    background-color: #f9f9f9;
  }
}

/* 图片网格样式 */
.image-grid {
  display: grid;
  grid-template-columns: repeat(3, 100px);
  gap: 16px;
  justify-content: center;
  margin-top: 40px;
}

.image-grid img {
  width: 100px;
  height: auto;
  border-radius: 8px;
  box-shadow: 0 2px 8px rgba(0, 0, 0, 0.1);
}
</style>
