<script setup lang="ts">
</script>

<template>
  <section class="max-w-3xl mx-auto px-4 py-12 text-gray-800">
    <h1 class="text-2xl font-bold mb-8">このツールについて</h1>

    <h2 class="text-xl font-semibold mb-4">背景</h2>
    <p class="mb-6 leading-relaxed">
      議事録をExcelで書く必要がある人にとって、書式の整備や決定事項・ToDoの抽出は面倒な作業です。
      このツールは、そうした課題を最小限の手間で解決するために設計されました。
    </p>

    <h2 class="text-xl font-semibold mb-4">主な特徴</h2>
    <ul class="list-disc list-inside mb-6 space-y-2">
      <li>完全ローカル処理（サーバへの通信ゼロ）</li>
      <li>決定事項・ToDoのExcel出力に特化</li>
      <li>書式整備された見やすい議事録レイアウト</li>
      <li>全ての操作はブラウザ内で完結</li>
    </ul>

    <h2 class="text-xl font-semibold mb-4">安心できる理由</h2>
    <ul class="list-disc list-inside mb-6 space-y-2">
      <li>通信なし。入力情報は一切外部に送信されません。</li>
      <li>データ保存もしておらず、ページを閉じると情報は消えます。</li>
      <li>オープンソースで開発され、コードは全てGitHubで公開済みです。</li>
    </ul>

    <h2 class="text-xl font-semibold mb-4">想定されるユーザー</h2>
    <ul class="list-disc list-inside mb-6 space-y-2">
      <li>社内の定例会議で記録を取る若手社員</li>
      <li>クライアントとの面談内容を記録する営業担当</li>
      <li>取締役会などの決定事項を文書化する管理部門</li>
    </ul>

    <h2 class="text-xl font-semibold mb-4">今後の展望</h2>
    <ul class="list-disc list-inside mb-6 space-y-2">
      <li>ユーザーフィードバックを受けてのUI等改善</li>
      <li>書式のカスタマイズ機能の追加（対外文書としての提出用フォーマットへの切替）</li>
    </ul>

    <h2 class="text-xl font-semibold mb-4">出力されるExcelのイメージ</h2>
    <img src="/excel_sample.png" alt="Excel出力例" class="rounded shadow max-w-full border" />
  </section>
</template>

<style scoped>
</style>
