<script setup lang="ts">
import { convertXlsxToIsolationDocx } from '#isolation'

const toast = useToast()

const file = ref<File | null>(null)
const loading = ref(false)
const lastMeta = ref<{ rows: number, nonCompliant: string } | null>(null)

watch(file, () => {
  lastMeta.value = null
})

async function convert() {
  if (!file.value) {
    toast.add({
      title: 'Выберите файл',
      description: 'Нужен Excel-шаблон (.xlsx)',
      color: 'warning',
      icon: 'i-lucide-triangle-alert'
    })
    return
  }

  loading.value = true
  try {
    const { blob, protocol, fileName } = await convertXlsxToIsolationDocx(file.value)
    const rows = protocol.rows.filter(r => r.kind === 'data').length
    const nonCompliant = protocol.nonCompliantIndexes.join(',')
    lastMeta.value = { rows, nonCompliant }

    const url = URL.createObjectURL(blob)
    const a = document.createElement('a')
    a.href = url
    a.download = fileName
    a.click()
    URL.revokeObjectURL(url)

    toast.add({
      title: 'Готово',
      description: `Сформирован протокол: ${rows} позиций`,
      color: 'success',
      icon: 'i-lucide-check'
    })
  } catch (err) {
    toast.add({
      title: 'Не удалось сформировать документ',
      description: err instanceof Error ? err.message : 'Неизвестная ошибка',
      color: 'error',
      icon: 'i-lucide-circle-x'
    })
  } finally {
    loading.value = false
  }
}
</script>

<template>
  <UContainer class="py-10 sm:py-16">
    <div class="mx-auto max-w-2xl space-y-8">
      <div class="space-y-3">
        <h1 class="text-3xl font-semibold tracking-tight text-highlighted sm:text-4xl">
          Протокол проверки изоляции
        </h1>
        <p class="text-base leading-relaxed text-muted">
          Загрузите Excel-шаблон со щитами — на выходе будет Word-таблица
          в альбомной ориентации по правилам протокола изоляции.
        </p>
      </div>

      <UCard>
        <template #header>
          <div class="flex items-center gap-2">
            <UIcon
              name="i-lucide-upload"
              class="size-5 text-primary"
            />
            <span class="font-medium">Загрузка шаблона</span>
          </div>
        </template>

        <div class="space-y-6">
          <UFileUpload
            v-model="file"
            accept=".xlsx,application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            label="Перетащите .xlsx или нажмите для выбора"
            description="Листы-щиты кроме sys, «Осн. инфо» и «ОБРАЗЕЦ»"
            class="min-h-40 w-full"
          />

          <div
            v-if="file"
            class="flex items-center gap-3 rounded-lg bg-elevated px-3 py-2 text-sm"
          >
            <UIcon
              name="i-lucide-file-spreadsheet"
              class="size-4 shrink-0"
            />
            <span class="truncate">{{ file.name }}</span>
            <span class="ms-auto shrink-0 text-muted">{{ (file.size / 1024).toFixed(1) }} КБ</span>
          </div>

          <UButton
            block
            size="lg"
            icon="i-lucide-file-down"
            :loading="loading"
            :disabled="!file"
            @click="convert"
          >
            Сформировать .docx
          </UButton>
        </div>
      </UCard>

      <UAlert
        v-if="lastMeta"
        color="success"
        variant="subtle"
        icon="i-lucide-circle-check"
        title="Документ скачан"
        :description="lastMeta.nonCompliant
          ? `Позиций: ${lastMeta.rows}. Не соответствуют: п/п ${lastMeta.nonCompliant}`
          : `Позиций: ${lastMeta.rows}. Все измерения соответствуют нормам.`"
      />

      <UCard variant="subtle">
        <div class="space-y-2 text-sm text-muted">
          <p class="font-medium text-highlighted">
            Как это работает
          </p>
          <ul class="list-disc space-y-1 ps-5">
            <li>Всё обрабатывается локально в браузере — файлы никуда не отправляются</li>
            <li>Читаются листы щитов и справочник диапазонов с листа <code>sys</code></li>
            <li>В конце добавляется заключение с перечнем несоответствий (если есть)</li>
          </ul>
        </div>
      </UCard>
    </div>
  </UContainer>
</template>
