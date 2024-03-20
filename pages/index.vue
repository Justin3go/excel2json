<template>
	<div>
		<v-file-input
			v-model="file"
			class="mx-12"
			chips
			label="File input"
			placeholder="Upload your documents"
			prepend-icon="mdi-paperclip"
		>
		</v-file-input>
    <div>{{ result }}</div>
	</div>
</template>

<script lang="ts" setup>
import Excel from "exceljs";

const file = ref();
const result = ref("No file selected");
const workbook = new Excel.Workbook();

const convertFormat = (data: any) => {
    const result = [];
    for (const key in data) {
        const model = data[key].model;
        const row = model.top - 1;
        const col = model.left - 1;
        const rowspan = model.bottom - model.top + 1;
        const colspan = model.right - model.left + 1;
        result.push({ row, col, rowspan, colspan });
    }
    // Sort the result by row and then by col
    result.sort((a, b) => a.row - b.row || a.col - b.col);
    return result;
};

watch(file, async (newVal) => {
	if (!newVal) {
    result.value = "No file selected";
    return;
  }
	console.log("file: ", newVal[0]);
	await workbook.xlsx.load(newVal[0]);
	workbook.eachSheet(function (worksheet, sheetId) {
    // @ts-ignore
    result.value = JSON.stringify(convertFormat(worksheet._merges));
	});
});
</script>

<style></style>
