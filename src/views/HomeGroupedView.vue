<script setup lang="ts">
import { ref, computed } from 'vue'
import * as ExcelJS from 'exceljs'
import { saveAs } from 'file-saver'
import type { RowValues } from 'exceljs'

function getJstDateString(): string {
  const jst = new Date(Date.now() + 9 * 60 * 60 * 1000) // JST = UTC+9
  return jst.toISOString().substring(0, 10)
}

function getRoundedTime(offsetMinutes: number = 0) {
  const now = new Date()
  now.setMinutes(now.getMinutes() + offsetMinutes)
  const m = now.getMinutes()
  const roundedMinutes = Math.ceil(m / 30) * 30
  if (roundedMinutes === 60) {
    now.setHours(now.getHours() + 1)
    now.setMinutes(0)
  } else {
    now.setMinutes(roundedMinutes)
  }
  return now.toTimeString().substring(0, 5)
}

const date = ref(getJstDateString())
const startTime = ref(getRoundedTime()) //30分単位の繰り上げ
const endTime = ref(getRoundedTime(60)) //その1時間後
const place = ref('')
const participantGroups = ref([{ company: '', members: [''] }])

const topics = ref([
  {
    title: '',
    memos: [''],
    annotations: [{ speaker: '', note: '' }],
    decisions: [{ content: '', person: '', deadline: '' }],
    todos: [{ content: '', person: '', deadline: '' }]
  }
])

const timeOptions = Array.from({ length: 24 * 4 }, (_, i) => {
  const h = Math.floor(i / 4).toString().padStart(2, '0')
  const m = (i % 4) * 15
  return `${h}:${m.toString().padStart(2, '0')}`
})

const participantNames = computed(() => participantGroups.value.flatMap(group => group.members))

function getDayOfWeekString(dateString: string): string {
  const days = ['日', '月', '火', '水', '木', '金', '土']
  return days[new Date(dateString).getDay()]
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
  const wb = new ExcelJS.Workbook()
  const buffer = await file.arrayBuffer()
  await wb.xlsx.load(buffer)
  const ws = wb.worksheets[0]
  const rows = ws.getSheetValues().slice(1)

  let i = 0
  const section = (label: string) => (rows[i]?.[1] || '').toString().trim().startsWith(label)
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

  topics.value = []
  while (advanceUntil('議題')) {
    const title = cell(rows[i], 1).split('：')[1] || ''
    i++
    const topic = {
      title,
      memos: [] as string[],
      annotations: [] as { speaker: string; note: string }[],
      decisions: [] as { content: string; person: string; deadline: string }[],
      todos: [] as { content: string; person: string; deadline: string }[]
    }

    if (advanceUntil('メモ')) {
      i++
      while (i < rows.length && rows[i]?.[1]) {
        topic.memos.push(cell(rows[i], 1))
        i++
      }
    }

    if (advanceUntil('発言')) {
      i++
      if (cell(rows[i], 1) === '発言者') i++
      while (i < rows.length && rows[i]?.[1]) {
        topic.annotations.push({ speaker: cell(rows[i], 1), note: cell(rows[i], 2) })
        i++
      }
    }

    if (advanceUntil('決定事項')) {
      i++
      if (cell(rows[i], 1) === '担当者') i++
      while (i < rows.length && rows[i]?.[1]) {
        topic.decisions.push({ person: cell(rows[i], 1), content: cell(rows[i], 2), deadline: cell(rows[i], 3).split('（')[0] })
        i++
      }
    }

    if (advanceUntil('ToDo')) {
      i++
      if (cell(rows[i], 1) === '担当者') i++
      while (i < rows.length && rows[i]?.[1]) {
        topic.todos.push({ person: cell(rows[i], 1), content: cell(rows[i], 2), deadline: cell(rows[i], 3).split('（')[0] })
        i++
      }
    }

    topics.value.push(topic)
  }
}

async function exportToExcel() {
  const wb = new ExcelJS.Workbook()
  const ws = wb.addWorksheet('議事録', {
    pageSetup: { paperSize: 9, orientation: 'portrait' },
    properties: { defaultRowHeight: 20 }
  })

  const boldStyle = { bold: true, size: 14 }
  const borderStyle: Partial<ExcelJS.Borders> = {
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
    row.eachCell(cell => {
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

  addRow(['日付', formatWithWeekday(date.value)], 'primary')
  ws.getRow(rowIndex - 1).getCell(2).fill = undefined
  addRow(['時間', `${startTime.value} 〜 ${endTime.value}`], 'primary')
  ws.getRow(rowIndex - 1).getCell(2).fill = undefined
  addRow(['場所', place.value], 'primary')
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

  topics.value.forEach((topic, index) => {
    addRow([`議題 ${index + 1}：${topic.title}`], 'primary', true, 2)

    addRow(['メモ'], 'primary')
    topic.memos.forEach((memo, idx) => {
      const style = idx % 2 === 1 ? 'even' : 'normal'
      addRow([memo], style, true, 2)
    })
    rowIndex++

    addRow(['発言'], 'primary')
    addRow(['発言者', '内容'], 'secondary')
    topic.annotations.forEach((a, idx) => {
      const style = idx % 2 === 1 ? 'even' : 'normal'
      addRow([a.speaker, a.note], style, true)
    })
    rowIndex++

    addRow(['決定事項'], 'primary')
    addRow(['担当者', '内容', '期日'], 'secondary')
    topic.decisions.forEach((d, idx) => {
      const style = idx % 2 === 1 ? 'even' : 'normal'
      addRow([d.person, d.content, formatWithWeekday(d.deadline)], style, true)
    })
    rowIndex++

    addRow(['ToDo'], 'primary')
    addRow(['担当者', '内容', '期日'], 'secondary')
    topic.todos.forEach((t, idx) => {
      const style = idx % 2 === 1 ? 'even' : 'normal'
      addRow([t.person, t.content, formatWithWeekday(t.deadline)], style, true)
    })
    rowIndex++
  })

  const buf = await wb.xlsx.writeBuffer()
  const blob = new Blob([buf], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' })
  saveAs(blob, `議事録_${date.value}.xlsx`)
}
</script>

<template>

  <p class="text-sm text-gray-500 mt-2">
  ※このアプリには保存機能がありません。ブラウザを更新・戻ると入力内容は消えます。<br/>
  ※このアプリは、議事録を作成するためのツールです。<br/>
  入力および取り込んだデータはブラウザ内部のみで扱い、外部には一切送信されません。<br/>
  ※このページは、議題に紐づけてメモなどを入力するためのページです。<br/>
  お好みに合わせてお使いください。なおExcel取込にもう片方との互換性はありません。<br/>
  <br/>
</p>

  <section class="max-w-4xl mx-auto px-4">
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
            <button @click="participantGroups.splice(i, 1)" class="text-base text-gray-500 hover:text-red-600">この会社を削除</button>
          </div>
        </div>
        <button @click="participantGroups.push({ company: '', members: [''] })" class="text-blue-600 text-base">＋ 会社グループを追加</button>
      </div>
    </div>

    <div class="mt-14">
      <h2 class="text-xl font-bold mb-6">議題ごとの内容</h2>
      <div class="space-y-10">
        <div v-for="(topic, i) in topics" :key="i" class="border p-6 rounded bg-white">
          <label class="text-base font-semibold mb-2 block">議題</label>
          <input v-model="topic.title" placeholder="議題を入力" class="w-full border rounded p-5 text-lg mb-6" />

          <h3 class="text-lg font-bold mb-3">メモ</h3>
          <div class="space-y-4 mb-6">
            <div v-for="(memo, j) in topic.memos" :key="j" class="flex gap-3">
              <textarea v-model="topic.memos[j]" placeholder="メモ内容" class="flex-1 border rounded p-5 text-lg min-h-[120px]" />
              <button @click="topic.memos.splice(j, 1)" class="text-red-500 text-base">✕</button>
            </div>
            <button @click="topic.memos.push('')" class="text-blue-600 text-base">＋ メモを追加</button>
          </div>

          <h3 class="text-lg font-bold mb-3">発言</h3>
          <div class="space-y-4 mb-6">
            <div v-for="(a, j) in topic.annotations" :key="j" class="flex flex-col md:flex-row gap-2">
              <select v-model="a.speaker" class="border p-2 rounded w-full md:w-1/3">
                <option value="">発言者</option>
                <option v-for="name in participantNames" :key="name" :value="name">{{ name }}</option>
              </select>
              <textarea v-model="a.note" placeholder="発言内容" class="border p-2 rounded w-full" rows="2" />
              <button @click="topic.annotations.splice(j, 1)" class="text-red-500">✕</button>
            </div>
            <button @click="topic.annotations.push({ speaker: '', note: '' })" class="text-blue-600 text-base">＋ 発言を追加</button>
          </div>

          <h3 class="text-lg font-bold mb-3">決定事項</h3>
          <div class="space-y-4 mb-6">
            <div v-for="(d, j) in topic.decisions" :key="j" class="flex flex-col md:flex-row gap-2">
              <textarea v-model="d.content" placeholder="内容" class="border p-4 rounded w-full md:w-1/2" />
              <input v-model="d.person" placeholder="担当者" class="border p-4 rounded w-full md:w-1/4" />
              <input v-model="d.deadline" type="date" class="border p-4 rounded w-full md:w-1/4" />
              <button @click="topic.decisions.splice(j, 1)" class="text-red-500">✕</button>
            </div>
            <button @click="topic.decisions.push({ content: '', person: '', deadline: '' })" class="text-blue-600 text-base">＋ 決定事項を追加</button>
          </div>

          <h3 class="text-lg font-bold mb-3">ToDo</h3>
          <div class="space-y-4">
            <div v-for="(t, j) in topic.todos" :key="j" class="flex flex-col md:flex-row gap-2">
              <textarea v-model="t.content" placeholder="内容" class="border p-4 rounded w-full md:w-1/2" />
              <input v-model="t.person" placeholder="担当者" class="border p-4 rounded w-full md:w-1/4" />
              <input v-model="t.deadline" type="date" class="border p-4 rounded w-full md:w-1/4" />
              <button @click="topic.todos.splice(j, 1)" class="text-red-500">✕</button>
            </div>
            <button @click="topic.todos.push({ content: '', person: '', deadline: '' })" class="text-blue-600 text-base">＋ ToDo を追加</button>
          </div>

          <div class="mt-6">
            <button @click="topics.splice(i, 1)" class="text-base text-gray-500 hover:text-red-600">この議題を削除</button>
          </div>
        </div>
        <button @click="topics.push({ title: '', memos: [''], annotations: [{ speaker: '', note: '' }], decisions: [{ content: '', person: '', deadline: '' }], todos: [{ content: '', person: '', deadline: '' }] })" class="text-blue-600 text-base">＋ 議題を追加</button>
      </div>
    </div>

    <div class="mt-14 text-right">
      <button @click="exportToExcel" class="bg-green-600 text-white text-base py-3 px-6 rounded hover:bg-green-700">
        Excel出力（議題ごと）
      </button>
    </div>

    <div class="my-4">
      <label class="font-bold block mb-2">Excelから読み込む（このページで作成したExcelのみ取込可能）</label>
      <input type="file" accept=".xlsx" @change="handleFileUpload" class="border p-2 rounded" />
    </div>
  </section>
</template>

<style scoped>
</style>
