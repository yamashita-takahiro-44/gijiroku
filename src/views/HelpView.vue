<script setup lang="ts">
</script>

<template>
  <section class="max-w-3xl mx-auto p-6 space-y-8 text-base leading-relaxed">
    <h1 class="text-2xl font-bold">このツールの使い方</h1>

    <p>
      このツールは、会議中にそのまま使える「軽量な議事録作成ツール」です。録音や自動要約ではなく、<strong>自分でその場で書きながらまとめる</strong>ことに特化しています。
    </p>

    <h2 class="text-xl font-semibold">基本の流れ</h2>
    <ol class="list-decimal list-inside space-y-2">
      <li>日付・時間・場所を入力</li>
      <li>参加者（会社名ごと）を追加</li>
      <li>議題を入力し、議題ごとに「メモ」「発言」「決定事項」「ToDo」を書く</li>
      <li>入力が終わったら、下部の「Excel出力」ボタンでファイルをダウンロード</li>
    </ol>

    <h2 class="text-xl font-semibold">発言・決定事項・ToDoの入力</h2>
    <ul class="list-disc list-inside space-y-2">
      <li><strong>発言</strong>は発言者名を選び、内容を記入します</li>
      <li><strong>決定事項・ToDo</strong>は「担当者」「内容」「期日」をそれぞれ記入します</li>
      <li>行は自由に追加・削除できます</li>
    </ul>

    <h2 class="text-xl font-semibold">Excelファイルの読み込み</h2>
    <p>
      過去に本ツールから出力したExcelファイルを読み込めば、内容を復元できます。<br>
      読み込むには「Excelから読み込む」欄でファイルを選択してください。
    </p>

    <h2 class="text-xl font-semibold">注意事項</h2>
    <ul class="list-disc list-inside space-y-2">
      <li><strong>保存機能はありません。</strong>ブラウザを更新・戻るなどすると内容は消えます</li>
      <li>出力されたExcelファイルのみが「保存」の役割を果たします</li>
      <li>このツールは <strong>通信を一切行わず、すべてローカルで完結</strong>します</li>
    </ul>

    <h2 class="text-xl font-semibold">推奨環境</h2>
    <p>
      パソコン（Chrome / Edge / Firefox）での使用を想定しています。<br>
      スマートフォンでも操作は可能ですが、画面が狭いため推奨はしていません。
    </p>

    <h2 class="text-xl font-semibold">要望・お問い合わせ</h2>
    <p>
      要望やお問い合わせのご希望がありましたら、Xアカウントへご連絡ください。<br>@yamashita_44<br>
    </p>
  </section>
</template>

<style scoped>
</style>
