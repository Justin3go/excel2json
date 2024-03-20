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
		<div>{{ mergeResult }}</div>
		<div>{{ tableResult }}</div>
	</div>
</template>

<script lang="ts" setup>
import Excel from "exceljs";

const file = ref();
const mergeResult = ref("No file selected");
const tableResult = ref("No file selected");
const workbook = new Excel.Workbook();

watch(file, async (newVal) => {
	if (!newVal) {
		mergeResult.value = "No file selected";
		return;
	}

	await workbook.xlsx.load(newVal[0]);
	workbook.eachSheet(function (worksheet, sheetId) {
		// @ts-ignore
		mergeResult.value = JSON.stringify(convertMergeFormat(worksheet._merges));
		// @ts-ignore
    tableResult.value = JSON.stringify(convertTableFormat(worksheet._rows));
	});
});

const convertMergeFormat = (data: any) => {
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

const convertTableFormat = (data: any) => {
	return data.map((item: any) => item.values);
};
</script>

<style></style>
