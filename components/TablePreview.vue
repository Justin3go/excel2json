<template>
	<div class="d-flex justify-center">
		<div>
			<div class="d-flex justify-center my-4">
				<v-btn class="text-capitalize" @click="scrollToPreview">Preview</v-btn>
			</div>
      <div style="min-height: 100vh;">
        <div class="d-flex justify-center">
          <v-switch v-model="isShowHeader" label="Show Header" inset></v-switch>
        </div>
        <div style="max-width: 100vw;">
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

const colLength = computed(() => {
	return props.tableData[0] ? Object.keys(props.tableData[0]).length : 0;
});
const headers = computed(() =>
	Array.from({ length: colLength.value - 1 }, (_, i) => `excel2json_key_${i + 1}`)
);
const isShowHeader = ref(false)

const scrollToPreview = () => {
	window.scrollTo({
		top: window.innerHeight + 20,
		behavior: "smooth",
	});
};
</script>

<style></style>
