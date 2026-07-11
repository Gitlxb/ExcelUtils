<template>
  <FunctionCard title="N层级文件夹文件提取" description="上传ZIP压缩包，提取其中所有文件到平级目录" :single-file="true" file-accept=".zip" :show-params="false" api-endpoint="/api/general/extract-files" />
</template>
<script>import FunctionCard from '@/components/FunctionCard.vue'; export default { name: 'NcengjiwenjjFzzt', components: { FunctionCard } }</script>
