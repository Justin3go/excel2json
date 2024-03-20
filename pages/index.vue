<template>
	<div>
		<header>
			<div class="text-h4 text-center my-6">
				Converting Excel to Front-end JSON Configuration Code (<a href="https://bug404.dev/demo.gif">Usage</a>)
			</div>
		</header>
		<v-file-input
			v-model="file"
			class="mx-12"
			chips
			label="upload your excel file"
			placeholder="Upload your documents"
			prepend-icon=""
			variant="outlined"
		>
		</v-file-input>
		<div class="d-flex">
			<div style="width: 50vw;">
				<div class="d-flex justify-space-between mb-1">
					<span class='text-h6'>Merge Cells</span>
          <v-snackbar :timeout="2000">
            <template #activator="{ props }">
              <v-btn v-bind="props" class="text-capitalize" variant="plain" color="primary" @click="copyText(JSON.stringify(mergeResult))">copy</v-btn>
            </template>
            copied to clipboard.
          </v-snackbar>
				</div>
				<v-sheet class="scrollable" height="calc(100vh - 250px)" width="100%" border>
					<vue-json-pretty
						:data="mergeResult"
						showIcon
						showLineNumber
						showLength
					></vue-json-pretty>
				</v-sheet>
			</div>
			<div style="width: 50vw;">
				<div class="d-flex justify-space-between mb-1">
					<span class='text-h6'>Table Data</span>
          <v-snackbar :timeout="2000">
            <template #activator="{ props }">
              <v-btn v-bind="props" class="text-capitalize" variant="plain" color="primary" @click="copyText(JSON.stringify(tableResult))">copy</v-btn>
            </template>
            copied to clipboard.
          </v-snackbar>
				</div>
				<v-sheet class="scrollable" height="calc(100vh - 250px)" width="100%" border>
					<vue-json-pretty
						:data="tableResult"
						showIcon
						showLineNumber
						showLength
					></vue-json-pretty>
				</v-sheet>
			</div>
		</div>
		<table-preview :mergeCells="mergeResult" :tableData="tableResult"></table-preview>
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

useSeoMeta({
	ogTitle: "Excel2json Tool",
	ogDescription: `Converting Excel to Front-end JSON Configuration Code`,
	ogImage: "https://bug404.dev/logo.jpg",
	twitterCard: "summary",
});

watch(file, async (newVal) => {
	if (!newVal && !newVal[0]) {
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
