package main

import (
	"fmt"
	"github.com/360EntSecGroup-Skylar/excelize"
	"strconv"
)

func main() {
	xlsx, err := excelize.OpenFile("/Users/rint/Documents/GolandProjects/mergeExcel/res/test.xlsx")
	if err != nil {
		fmt.Println(err)
		return
	}
	// 获取工作表中指定单元格的值
	//cell := xlsx.GetCellValue("1", "B2")
	//fmt.Println(cell)
	fmt.Println("sheet counts:",xlsx.SheetCount);
	for i := 1; i <= 2; i++ {
		sheetName := xlsx.GetSheetName(i)
		mergeCells := xlsx.GetMergeCells(sheetName)
		for _, mergeCell := range mergeCells {
			println(mergeCell.GetStartAxis() + "->" + mergeCell.GetEndAxis() + ", v:" + mergeCell.GetCellValue());
		}
		for _, row := range xlsx.Sheet["xl/worksheets/sheet"+strconv.Itoa(i)+".xml"].SheetData.Row {
			for _, cell := range row.C {
				cellValue := xlsx.GetCellValue(sheetName, cell.R)
				fmt.Print(cellValue, "\t");
			}
			fmt.Println()
		}

	}

	for k,_ := range xlsx.Sheet {
		println(k);
	}
	//for i:=1; i<=2; i++ {
	//	rows := xlsx.GetRows(xlsx.GetSheetName(i))
	//	for _, row := range rows {
	//		for _, colCell := range row {
	//			fmt.Print(colCell, "\t")
	//		}
	//		fmt.Println()
	//	}
	//}

	//for i:=1; i<=2; i++ {
	//	rows := xlsx.GetRows(xlsx.GetSheetName(i))
	//	for _, row := range rows {
	//		for _, colCell := range row {
	//			fmt.Print(colCell, "\t")
	//		}
	//		fmt.Println()
	//	}
	//}
	fmt.Println("end ...");
}
