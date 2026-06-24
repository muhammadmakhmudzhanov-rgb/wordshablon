<script setup>
import { ref, onMounted } from 'vue'
import PizZip from 'pizzip'
import Docxtemplater from 'docxtemplater'
import { saveAs } from 'file-saver'
import { renderAsync } from 'docx-preview'

const surname = ref('')
const name = ref('')
const patronymic = ref('')
const citizenship = ref('')
const isLoaded = ref(false)
const preview = ref(null)


const fields = ref([])
const formData = ref({})

let templateBuffer = null

const updatePreview = async () => {
  if (!templateBuffer || !preview.value) return

  const zip = new PizZip(templateBuffer)

  const doc = new Docxtemplater(zip)

  doc.render(formData.value)

  const blob = doc.getZip().generate({
    type: 'blob',
    mimeType:
      'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
  })

  preview.value.innerHTML = ''

  await renderAsync(blob, preview.value)

  preview.value.contentEditable = true
}

const generateDoc = async () => {
  const zip = new PizZip(templateBuffer)

  const doc = new Docxtemplater(zip)

  doc.render(formData.value)

  const blob = doc.getZip().generate({
    type: 'blob',
    mimeType:
      'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
  })

  saveAs(blob, 'document.docx')
}
const uploadTemplate = async (event) => {
  const file = event.target.files[0]

  if (!file) return

  templateBuffer = await file.arrayBuffer()

  const zip = new PizZip(templateBuffer)

  const xml = zip.files['word/document.xml'].asText()

  const matches = [
    ...new Set(
      [...xml.matchAll(/\{([^}]+)\}/g)].map(
        match => match[1]
      )
    ),
  ]

  fields.value = matches

  formData.value = {}

  matches.forEach(field => {
    formData.value[field] = ''
  })

  isLoaded.value = true

  await updatePreview()
}
</script>

<template>
<div v-if="!isLoaded" class="empty">
  <div class="upload-card">
    <h1>📄 Конструктор документов</h1>

    <p>
      Загрузите Word шаблон (.docx)
    </p>

    <label class="upload-btn">
      Выбрать файл
      <input
        type="file"
        accept=".docx"
        @change="uploadTemplate"
      >
    </label>
  </div>
</div>

<div v-else class="wrapper">
  <button class="save-btn" @click="generateDoc">
    Скачать Word
  </button>

  <div
    ref="preview"
    class="preview"
    contenteditable="true"
  ></div>
</div>
</template>

<style>
* {
  margin: 0;
  padding: 0;
  box-sizing: border-box;
}

body {
  background: #eef2f7;
  font-family: Inter, Arial, sans-serif;
}

.wrapper {
  display: flex;
  gap: 30px;
  padding: 30px;
  min-height: 100vh;
}

.form {
  width: 350px;
  background: white;
  padding: 25px;
  border-radius: 20px;
  box-shadow: 0 10px 30px rgba(0, 0, 0, .08);
  height: fit-content;
  position: sticky;
  top: 20px;
}

.form h2 {
  margin-bottom: 20px;
  color: #1e293b;
}

.form input {
  width: 100%;
  padding: 14px;
  border: 2px solid #e2e8f0;
  border-radius: 12px;
  margin-bottom: 12px;
  font-size: 15px;
  transition: .3s;
}

.form input:focus {
  outline: none;
  border-color: #3b82f6;
}

.form button {
  width: 100%;
  padding: 14px;
  border: none;
  border-radius: 12px;
  font-size: 15px;
  cursor: pointer;
  background: linear-gradient(135deg,
      #3b82f6,
      #2563eb);
  color: white;
  transition: .3s;
}

.form button:hover {
  transform: translateY(-2px);
}

.preview {
  flex: 1;
  background: white;
  border-radius: 20px;
  padding: 40px;
  box-shadow: 0 10px 30px rgba(0, 0, 0, .08);
  overflow: auto;
  min-height: 90vh;
}

/* Стили для отображаемого Word */

.preview section {
  width: 794px;
  min-height: 1123px;

  margin: auto;

  background: white;

  padding: 60px;

  box-shadow:
    0 0 20px rgba(0,0,0,.2);

  border: 1px solid #e5e7eb;
}

.preview table {
  border-collapse: collapse;
}

.preview td,
.preview th {
  border: 1px solid #dbe2ea;
}

.empty {
  height: 100vh;

  display: flex;
  justify-content: center;
  align-items: center;

  background: linear-gradient(
    135deg,
    #eef2ff,
    #f8fafc
  );
}

.upload-card {
  width: 500px;

  background: white;

  padding: 50px;

  border-radius: 24px;

  text-align: center;

  box-shadow:
    0 20px 40px rgba(0,0,0,.08);
}

.upload-card h1 {
  margin-bottom: 15px;

  color: #1e293b;
}

.upload-card p {
  color: #64748b;

  margin-bottom: 30px;
}

.upload-btn {
  display: inline-block;

  padding: 15px 35px;

  background: #2563eb;

  color: white;

  border-radius: 12px;

  cursor: pointer;

  transition: .3s;
}

.upload-btn:hover {
  transform: translateY(-3px);
}

.upload-btn input {
  display: none;
}

.wrapper {
  padding: 30px;
}

.save-btn {
  position: fixed;
  top: 20px;
  right: 20px;

  padding: 12px 20px;

  border: none;
  border-radius: 10px;

  background: #2563eb;
  color: white;

  cursor: pointer;
  z-index: 999;
}
</style>