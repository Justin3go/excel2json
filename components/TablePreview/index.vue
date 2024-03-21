<template>
	<div class="d-flex justify-center">
		<div>
			<div class="d-flex justify-center my-4">
				<v-btn class="text-capitalize" @click="scrollToPreview">Preview</v-btn>
			</div>
			<div style="min-height: 100vh">
				<div class="d-flex justify-center align-center">
					<v-btn
						variant="plain"
						class="text-capitalize mx-2"
						style="margin-bottom: 22px"
						@click="exportPDF"
						>Export PDF</v-btn
					>
					<v-btn
						variant="plain"
						class="text-capitalize ml-2 mr-6"
						style="margin-bottom: 22px"
						@click="exportExcel"
						>Export Excel</v-btn
					>
					<v-switch v-model="isShowHeader" label="Show Header" inset></v-switch>
				</div>
				<div id="table-id" style="max-width: 100vw">
					<vxe-table
						border
						show-footer
						ref="tableRef"
						align="center"
						empty-text="no data"
						:showHeader="isShowHeader"
						:print-config="{}"
						:export-config="{}"
						:column-config="{ resizable: true, width: 90 }"
						:merge-cells="props.mergeCells"
						:data="props.tableData"
					>
						<vxe-column v-for="col in headers" :field="col" :title="col" />
					</vxe-table>
				</div>
			</div>
		</div>
	</div>
</template>

<script lang="ts" setup>
interface IProps {
	mergeCells: {
		row: number;
		col: number;
		rowspan: number;
		colspan: number;
	}[];
	tableData: any;
}

const props = withDefaults(defineProps<IProps>(), {
	mergeCells: [] as any,
	tableData: [],
});

const tableRef = ref();

const colLength = computed(() => {
	return props.tableData[0] ? Object.keys(props.tableData[0]).length : 0;
});
const headers = computed(() =>
	Array.from(
		{ length: colLength.value - 1 },
		(_, i) => `excel2json_key_${i + 1}`
	)
);
const isShowHeader = ref(false);

const scrollToPreview = () => {
	window.scrollTo({
		top: window.innerHeight + 20,
		behavior: "smooth",
	});
};

const exportPDF = () => {
	htmlToPDF("table-id", "demoName");
};

const exportExcel = () => {
	tableRef.value.exportData({
		filename: new Date().toISOString(),
		sheetName: "sheet1",
		type: "xlsx",
		useStyle: true,
		isFooter: false,
		isHeader: false,
		download: true,
		sheetMethod: ({ worksheet }: any) => {
			props.mergeCells.forEach((item) => {
				const { row, col, rowspan, colspan } = item;
				console.log("curMerge: ", item);
				worksheet.mergeCells(row + 1, col + 1, row + rowspan, col + colspan); // 可能需要最后两个减一
			});
			// 导出样式demo
			// worksheet.getCell("A1").fill = {
			// 	type: "pattern",
			// 	pattern: "darkVertical",
			// 	fgColor: { argb: "FFFF0000" },
			// };
		},
	});
};
</script>

<style></style>
