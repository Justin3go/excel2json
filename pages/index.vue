<template>
	<div>
		<v-file-input
			v-model="file"
			class="mt-8 mx-12"
			chips
			label="upload your excel file"
			placeholder="Upload your documents"
			prepend-icon=""
			variant="outlined"
		>
		</v-file-input>
		<div class="d-flex">
			<v-sheet class="scrollable" height="80vh" width="50vw" border>
				<vue-json-pretty :data="mergeResult" showIcon showLineNumber showLength></vue-json-pretty>
			</v-sheet>
			<v-sheet class="scrollable" height="80vh" width="50vw" border>
				<vue-json-pretty :data="tableResult" showIcon showLineNumber showLength></vue-json-pretty>
			</v-sheet>
		</div>
	</div>
</template>

<script lang="ts" setup>
import Excel from "exceljs";
import VueJsonPretty from "vue-json-pretty";
import "vue-json-pretty/lib/styles.css";

const file = ref();
const mergeResult = ref();
const tableResult = ref();
const workbook = new Excel.Workbook();

watch(file, async (newVal) => {
	if (!newVal) {
		mergeResult.value = null;
		return;
	}

	await workbook.xlsx.load(newVal[0]);
	workbook.eachSheet(function (worksheet) {
		// @ts-ignore
		mergeResult.value = convertMergeFormat(worksheet._merges);
		// @ts-ignore
		tableResult.value = convertTableFormat(worksheet._rows);
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
	return data.map((item: any) =>
		Object.fromEntries(
			item.values
				.map((value: any, index: number) => [`excel2json_key_${index}`, value])
				.filter(Boolean)
		)
	);
};
</script>

<style scoped>
.scrollable {
	overflow: auto;
}
</style>
