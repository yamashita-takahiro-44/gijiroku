<script setup lang="ts">
import { ref, computed } from 'vue'
import * as ExcelJS from 'exceljs'
import { saveAs } from 'file-saver'
import type { RowValues } from 'exceljs'

function getJstDateString(): string {
  const jst = new Date(Date.now() + 9 * 60 * 60 * 1000) // JST = UTC+9
  return jst.toISOString().substring(0, 10)
}

const date = ref(getJstDateString())
const startTime = ref('09:00')
const endTime = ref('10:00')
const place = ref('')
const topics = ref<string[]>([''])
const participantGroups = ref([
  {
    company: '',
    members: ['']
  }
])

const memoItems = ref<string[]>([''])
const decisions = ref([{ content: '', person: '', deadline: '' }])
const todos = ref([{ content: '', person: '', deadline: '' }])
const isImporting = ref(false)

const annotations = ref([{ speaker: '', note: '' }])

const participantNames = computed(() => {
  return participantGroups.value.flatMap(group => group.members)
})

const timeOptions = Array.from({ length: 24 * 4 }, (_, i) => {
  const h = Math.floor(i / 4).toString().padStart(2, '0')
  const m = (i % 4) * 15
  const mm = m.toString().padStart(2, '0')
  return `${h}:${mm}`
})

function getDayOfWeekString(dateString: string): string {
  const days = ['日', '月', '火', '水', '木', '金', '土']
  const dateObj = new Date(dateString)
  const day = dateObj.getDay()
  return days[day]
}

function formatWithWeekday(dateString: string): string {
  if (!dateString) return ''
  return `${dateString}（${getDayOfWeekString(dateString)}）`
}

function handleFileUpload(event: Event) {
  const target = event.target as HTMLInputElement
  const file = target.files?.[0]
  if (file) {
    importFromExcel(file)
  }
}

async function importFromExcel(file: File) {
  isImporting.value = true
  const wb = new ExcelJS.Workbook()
  const buffer = await file.arrayBuffer()
  await wb.xlsx.load(buffer)
  const ws = wb.worksheets[0]
  const rows = ws.getSheetValues().slice(1) // 0-index is null

  let i = 0
  const section = (label: string) => (rows[i]?.[1] || '').toString().trim() === label
  const cell = (row: RowValues, index: number): string => {
  if (Array.isArray(row)) {
    return (row[index] ?? '').toString().trim()
  }
  return ''
}

  function advanceUntil(label: string) {
    while (i < rows.length && !section(label)) i++
    return i < rows.length
  }

  function collectList(
    startLabel: string,
    mapFn: (r: RowValues) => void,
    headerCheck: (r: RowValues) => boolean
  ) {
    if (!advanceUntil(startLabel)) return
    i++
    if (headerCheck(rows[i])) i++
    while (i < rows.length && Array.isArray(rows[i]) && rows[i]?.[1]) {
      const row = rows[i]
      mapFn(row)
      i++
    }
  }

  if (advanceUntil('日付')) {
    date.value = cell(rows[i], 2).split('（')[0] || ''
    i++
  }
  if (advanceUntil('時間')) {
    const time = cell(rows[i], 2).split('〜').map(s => s.trim())
    startTime.value = time[0] || ''
    endTime.value = time[1] || ''
    i++
  }
  if (advanceUntil('場所')) {
    place.value = cell(rows[i], 2)
    i++
  }

  participantGroups.value = []
  if (advanceUntil('参加者')) {
    i++
    if (cell(rows[i], 1) === '会社名') i++
    while (i < rows.length && rows[i]?.[1]) {
      const company = cell(rows[i], 1)
      const member = cell(rows[i], 2)
      let group = participantGroups.value.find(g => g.company === company)
      if (!group) {
        group = { company, members: [] }
        participantGroups.value.push(group)
      }
      group.members.push(member)
      i++
    }
  }

  const extractSimpleList = (label: string, target: typeof memoItems | typeof topics) => {
    target.value = []
    if (advanceUntil(label)) {
      i++
      while (i < rows.length && rows[i]?.[1]) {
        target.value.push(cell(rows[i], 1))
        i++
      }
    }
  }

  extractSimpleList('議題', topics)
  extractSimpleList('メモ', memoItems)

  annotations.value = []
  collectList('発言', (r) => {
    annotations.value.push({ speaker: cell(r, 1), note: cell(r, 2) })
  }, (r) => cell(r, 1) === '発言者')

  decisions.value.splice(0)
  collectList('決定事項', (r) => {
    decisions.value.push({ content: cell(r, 2), person: cell(r, 1), deadline: cell(r, 3).split('（')[0] || '' })
  }, (r) => cell(r, 1) === '担当者')

  todos.value.splice(0)
  collectList('ToDo', (r) => {
    todos.value.push({ content: cell(r, 2), person: cell(r, 1), deadline: cell(r, 3).split('（')[0] || '' })
  }, (r) => cell(r, 1) === '担当者')

  isImporting.value = false
}

async function exportToExcel() {
  const wb = new ExcelJS.Workbook()
  const ws = wb.addWorksheet('議事録', {
    pageSetup: { paperSize: 9, orientation: 'portrait' },
    properties: { defaultRowHeight: 20 }
  })

  const boldStyle = { bold: true, size: 14 }
  const borderStyle = {
    top: { style: 'thin' as ExcelJS.BorderStyle },
    bottom: { style: 'thin' as ExcelJS.BorderStyle },
    left: { style: 'thin' as ExcelJS.BorderStyle },
    right: { style: 'thin' as ExcelJS.BorderStyle }
  }

  const colors = {
    primary: 'D9D9D9',
    secondary: 'EEEEEE',
    evenRow: 'F7F7F7'
  }

  let rowIndex = 1

  function addRow(values: (string | number)[], style: 'primary' | 'secondary' | 'normal' | 'even' = 'normal', wrap = false, mergeAcross = 0) {
    const row = ws.insertRow(rowIndex++, values)
    row.eachCell((cell) => {
      cell.border = borderStyle
      cell.alignment = { vertical: 'top', wrapText: wrap }
      if (style === 'primary') {
        cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: colors.primary } }
        cell.font = boldStyle
      } else if (style === 'secondary') {
        cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: colors.secondary } }
        cell.font = { bold: true }
      } else if (style === 'even') {
        cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: colors.evenRow } }
      }
    })
    if (mergeAcross > 0) {
      ws.mergeCells(row.number, 1, row.number, mergeAcross + 1)
    }
    return row
  }

  ws.columns = [
    { width: 20 },
    { width: 60 },
    { width: 60 }
  ]

  const formattedDate = formatWithWeekday(date.value)

  addRow(['日付', formattedDate], 'primary', false);
  ws.getRow(rowIndex - 1).getCell(2).fill = undefined
  addRow(['時間', `${startTime.value} 〜 ${endTime.value}`], 'primary', false);
  ws.getRow(rowIndex - 1).getCell(2).fill = undefined
  addRow(['場所', place.value], 'primary', false);
  ws.getRow(rowIndex - 1).getCell(2).fill = undefined
  rowIndex++

  addRow(['参加者'], 'primary')
  addRow(['会社名', '参加者名'], 'secondary')
  participantGroups.value.forEach(group => {
    group.members.forEach((member, idx) => {
      const style = idx % 2 === 1 ? 'even' : 'normal'
      addRow([group.company, member], style)
    })
  })
  rowIndex++

  addRow(['議題'], 'primary')
  topics.value.forEach((t, idx) => {
    const style = idx % 2 === 1 ? 'even' : 'normal'
    addRow([t], style, true, 2)
  })
  rowIndex++

  addRow(['メモ'], 'primary')
  memoItems.value.forEach((m, idx) => {
    const style = idx % 2 === 1 ? 'even' : 'normal'
    addRow([m], style, true, 2)
  })
  rowIndex++

  addRow(['発言'], 'primary')
  addRow(['発言者', '内容'], 'secondary')
  annotations.value.forEach((a, idx) => {
    const style = idx % 2 === 1 ? 'even' : 'normal'
    addRow([a.speaker, a.note], style, true)
  })
  rowIndex++

  addRow(['決定事項'], 'primary')
  addRow(['担当者', '内容', '期日'], 'secondary')
  decisions.value.forEach((d, idx) => {
    const style = idx % 2 === 1 ? 'even' : 'normal'
    addRow([d.person, d.content, formatWithWeekday(d.deadline)], style, true)
  })
  rowIndex++

  addRow(['ToDo'], 'primary')
  addRow(['担当者', '内容', '期日'], 'secondary')
  todos.value.forEach((t, idx) => {
    const style = idx % 2 === 1 ? 'even' : 'normal'
    addRow([t.person, t.content, formatWithWeekday(t.deadline)], style, true)
  })

  const buf = await wb.xlsx.writeBuffer()
  const blob = new Blob([buf], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' })
  saveAs(blob, `議事録_${date.value}.xlsx`)
}
</script>

<template>
  <p class="text-sm text-gray-500 mt-2">
    ※このアプリには保存機能がありません。ブラウザを更新・戻ると入力内容は消えます。<br />
    ※このアプリは、議事録を作成するためのツールです。<br />
    入力および取り込んだデータはブラウザ内部のみで扱い、外部には一切送信されません。<br /><br />
  </p>

  <section class="max-w-4xl mx-auto px-4">
    <!-- 日時・時間・場所 -->
    <div class="grid grid-cols-1 sm:grid-cols-3 gap-8">
      <div class="sm:col-span-1">
        <label class="block text-base font-semibold mb-3">日付</label>
        <input type="date" v-model="date" class="w-full border rounded p-5 text-lg" />
      </div>
      <div class="sm:col-span-2">
        <label class="block text-base font-semibold mb-3">開始〜終了時間</label>
        <div class="flex gap-3">
          <select v-model="startTime" class="w-1/2 border rounded p-6 text-lg">
            <option v-for="t in timeOptions" :key="'s' + t" :value="t">{{ t }}</option>
          </select>
          <span class="self-center text-base text-gray-500">〜</span>
          <select v-model="endTime" class="w-1/2 border rounded p-6 text-lg">
            <option v-for="t in timeOptions" :key="'e' + t" :value="t">{{ t }}</option>
          </select>
        </div>
      </div>
      <div class="sm:col-span-3">
        <label class="block text-base font-semibold mb-3">場所</label>
        <input type="text" v-model="place" placeholder="例：Zoom / 会議室A" class="w-full border rounded p-5 text-lg" />
      </div>
    </div>

    <!-- 参加者 -->
    <div class="mt-14">
      <h2 class="text-xl font-bold mb-6">参加者</h2>
      <div class="space-y-8">
        <div v-for="(group, i) in participantGroups" :key="i" class="border p-6 rounded bg-white">
          <label class="text-base font-semibold mb-3 block">会社名</label>
          <input v-model="group.company" placeholder="例：株式会社〇〇" class="w-full border rounded p-5 text-lg mb-5" />

          <div class="space-y-4">
            <div v-for="(member, j) in group.members" :key="j" class="flex gap-3">
              <input v-model="group.members[j]" placeholder="参加者名" class="flex-1 border rounded p-5 text-lg" />
              <button @click="group.members.splice(j, 1)" class="text-red-500 text-base">✕</button>
            </div>
            <button @click="group.members.push('')" class="text-blue-600 text-base">＋ メンバーを追加</button>
          </div>

          <div class="mt-4">
            <button @click="participantGroups.splice(i, 1)"
              class="text-base text-gray-500 hover:text-red-600">この会社を削除</button>
          </div>
        </div>
        <button @click="participantGroups.push({ company: '', members: [''] })" class="text-blue-600 text-base">＋
          会社グループを追加</button>
      </div>
    </div>

    <!-- 議題 -->
    <div class="mt-14">
      <h2 class="text-xl font-bold mb-6">議題</h2>
      <div class="space-y-4">
        <div v-for="(topic, i) in topics" :key="i" class="flex gap-3">
          <textarea v-model="topics[i]" placeholder="議題を入力" class="flex-1 border rounded p-5 text-lg" />
          <button @click="topics.splice(i, 1)" class="text-red-500 text-base">✕</button>
        </div>
        <button @click="topics.push('')" class="text-blue-600 text-base">＋ 議題を追加</button>
      </div>
    </div>

    <!-- メモ欄 -->
    <div class="mt-14">
      <h2 class="text-xl font-bold mb-6">メモ</h2>
      <div class="space-y-4">
        <div v-for="(memo, i) in memoItems" :key="i" class="flex gap-3">
          <textarea v-model="memoItems[i]" placeholder="メモ内容"
            class="flex-1 border rounded p-5 text-lg min-h-[120px]"></textarea>
          <button @click="memoItems.splice(i, 1)" class="text-red-500 text-base">✕</button>
        </div>
        <button @click="memoItems.push('')" class="text-blue-600 text-base">＋ メモを追加</button>
      </div>
    </div>

    <h2 class="text-lg font-bold mt-6 mb-2">発言</h2>
    <div v-for="(a, i) in annotations" :key="i" class="flex flex-col md:flex-row gap-2 mb-2">
      <select v-model="a.speaker" class="border p-2 rounded w-full md:w-1/3">
        <option value="">発言者</option>
        <option v-for="name in participantNames" :key="name" :value="name">{{ name }}</option>
      </select>
      <textarea v-model="a.note" placeholder="発言内容" class="border p-2 rounded w-full" rows="2" />
      <button @click="annotations.splice(i, 1)" class="text-red-500">✕</button>
    </div>
    <button @click="annotations.push({ speaker: '', note: '' })" class="text-blue-600 text-base">＋ 発言を追加</button>


    <!-- 決定事項 -->
    <div class="mt-14">
      <h2 class="text-lg font-bold mt-6 mb-2">決定事項</h2>
      <div v-for="(d, i) in decisions" :key="i" class="flex flex-col md:flex-row gap-2 mb-2">
        <textarea v-model="d.content" placeholder="内容" class="border p-4 rounded w-full md:w-1/2" />
        <input v-model="d.person" placeholder="担当者" class="border p-4 rounded w-full md:w-1/4" />
        <input v-model="d.deadline" type="date" class="border p-4 rounded w-full md:w-1/4" />
        <button @click="decisions.splice(i, 1)" class="text-red-500">✕</button>
      </div>
      <button @click="decisions.push({ content: '', person: '', deadline: '' })"
        class="text-blue-600 text-base">＋ 決定事項を追加</button>
    </div>

    <!-- ToDo -->
    <div class="mt-14">
      <h2 class="text-xl font-bold mb-6">ToDo</h2>
      <div class="space-y-4">
        <div v-for="(item, i) in todos" :key="i" class="flex flex-col md:flex-row gap-2 mb-2">
          <textarea v-model="item.content" placeholder="内容" class="border p-4 rounded w-full md:w-1/2" />
          <input v-model="item.person" placeholder="担当者" class="border p-4 rounded w-full md:w-1/4" />
          <input v-model="item.deadline" type="date" class="border p-4 rounded w-full md:w-1/4" />
          <button @click="todos.splice(i, 1)" class="text-red-500">✕</button>
        </div>
        <button @click="todos.push({ content: '', person: '', deadline: '' })" class="text-blue-600 text-base">＋ ToDo
          を追加</button>
      </div>
    </div>

    <!-- Excel出力 -->
    <div class="mt-14 text-right">
      <button @click="exportToExcel" class="bg-green-600 text-white text-base py-3 px-6 rounded hover:bg-green-700">
        Excel出力
      </button>
    </div>

    <div class="my-4">
      <label class="font-bold block mb-2">Excelから読み込む（このページで作成したExcelのみ取込可能）</label>
      <input type="file" accept=".xlsx" @change="handleFileUpload" class="border p-2 rounded" />
    </div>

  </section>
</template>

<style scoped></style>
