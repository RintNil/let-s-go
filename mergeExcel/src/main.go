package main

import (
	"fmt"
	"github.com/360EntSecGroup-Skylar/excelize"
	"io/ioutil"
	"os"
	"strconv"
)

func main() {

	wxlsx := excelize.NewFile()
	cstyle, cerr := wxlsx.NewStyle(`{
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
	sep := string(os.PathSeparator)
	dirFiles, readDirErr := ioutil.ReadDir("." + sep + "inputs")

	if readDirErr != nil {
		fmt.Println(readDirErr)
		fmt.Println("请输入正确的文件目录")
		return;
	}

	var toMergeFiles [] string
	for _, dirFile := range dirFiles {
		if dirFile.IsDir() {
			fmt.Println(dirFile.Name(), " is a dir, temp not support child dir ")
			continue
		} else {
			toMergeFiles = append(toMergeFiles, dirFile.Name())
		}
	}
	//sort.Strings(toMergeFiles)

	wxlsx.NewSheet("Sheet1");
	wxlsx.NewSheet("Sheet2");
	var rowIndexMap map[string]int = make(map[string]int, 2)
	rowIndexMap["Sheet1"] = 0
	rowIndexMap["Sheet2"] = 0
	for toMergeIndex, toMergeFile := range toMergeFiles {
		fmt.Println("开始 合并 "+toMergeFile+" ", toMergeIndex+1, "/", len(toMergeFiles))
		xlsx, err := excelize.OpenFile("." + sep + "inputs" + sep + toMergeFile)
		if err != nil {
			fmt.Println(toMergeFile+" 合并失败 cause ", err)
			continue
		}
		//fmt.Println("sheet counts:", xlsx.SheetCount)
		for i := 1; i <= 2; i++ {
			wSheetName := "Sheet" + strconv.Itoa(i);
			sRowIndex := rowIndexMap[wSheetName];
			rowIndex := rowIndexMap[wSheetName];
			sheetName := xlsx.GetSheetName(i)
			mergeCells := xlsx.GetMergeCells(sheetName)
			for _, row := range xlsx.Sheet["xl/worksheets/sheet"+strconv.Itoa(i)+".xml"].SheetData.Row {
				emptyNum := 0
				for _, cell := range row.C {
					cellValue := xlsx.GetCellValue(sheetName, cell.R)
					cIndex, rIndex := separate(cell.R)
					cellIndex := cIndex + strconv.Itoa(rIndex+sRowIndex)
					wxlsx.SetCellValue(wSheetName, cellIndex, cellValue)
					wxlsx.SetCellStyle(wSheetName, cellIndex, cellIndex, cstyle)
					if (cellValue == "") {
						emptyNum++
					}
				}
				if (emptyNum == len(row.C)) {
					break
				}
				rowIndex++;
			}
			rowIndex = rowIndex + 2
			rowIndexMap[wSheetName] = rowIndex
			for _, mergeCell := range mergeCells {
				cSIndex, rSIndex := separate(mergeCell.GetStartAxis())
				cellSIndex := cSIndex + strconv.Itoa(rSIndex+sRowIndex)
				cEIndex, rEIndex := separate(mergeCell.GetEndAxis())
				cellEIndex := cEIndex + strconv.Itoa(rEIndex+sRowIndex)
				wxlsx.MergeCell(wSheetName, cellSIndex, cellEIndex);
			}
		}
		fmt.Println("合并 成功 "+toMergeFile, toMergeIndex+1, "/", len(toMergeFiles))
	}
	// 根据指定路径保存文件
	werr := wxlsx.SaveAs("." + sep + "Book1.xlsx")
	if werr != nil {
		fmt.Println(werr)
	}
	fmt.Println("end ...");
}

func separate(word string) (string, int) {
	wi := 0
	for _, w := range word {
		if (('a' <= w && w <= 'z') || ('A' <= w && w <= 'Z')) {
			wi++
			continue
		}
		break
	}
	num, _ := strconv.Atoi(word[wi:len(word)])
	return word[0:wi], num;
}
