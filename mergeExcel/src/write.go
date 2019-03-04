package main

import (
	"fmt"

	"github.com/360EntSecGroup-Skylar/excelize"
)

func main() {
	xlsx := excelize.NewFile()
	// 创建一个工作表
	index := xlsx.NewSheet("Sheet1")
	// 设置单元格的值
	xlsx.SetCellValue("Sheet1", "B2", "Hello world.")
	xlsx.SetCellValue("Sheet1", "A1", "Hello world.")
	xlsx.MergeCell("Sheet1", "A1", "B2")
	style, cerr := xlsx.NewStyle(`{
  "border" : [
    {
      "type" : "left",
      "color" : "000000",
      "style" : 1
    },
    {
      "type" : "top",
      "color" : "000000",
      "style" : 1
    },
    {
      "type" : "bottom",
      "color" : "000000",
      "style" : 1
    },
    {
      "type" : "right",
      "color" : "000000",
      "style" : 1
    }
  ],
  "Alignment" : {
    "horizontal" : "center",
    "vertical" : "center"
  }
}`)
	if cerr != nil {
		fmt.Println(cerr)
	}
	xlsx.SetCellStyle("Sheet1", "A1", "B2", style)
	// 设置工作簿的默认工作表
	xlsx.SetActiveSheet(index)
	// 根据指定路径保存文件
	err := xlsx.SaveAs("./Book1.xlsx")
	if err != nil {
		fmt.Println(err)
	}

}
