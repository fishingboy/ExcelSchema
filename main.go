package main

import (
	"fmt"
	"log"

	"github.com/xuri/excelize/v2"
)

//TIP <p>To run your code, right-click the code and select <b>Run</b>.</p> <p>Alternatively, click
// the <icon src="AllIcons.Actions.Execute"/> icon in the gutter and select the <b>Run</b> menu item from here.</p>

func main() {
	f, err := excelize.OpenFile("stage.xlsx")
	if err != nil {
		log.Fatal(err)
	}
	defer f.Close()

	// 取得第一個工作表名稱
	sheetName := f.GetSheetName(0)

	// 讀取指定工作表所有資料
	rows, err := f.GetRows(sheetName)
	if err != nil {
		log.Fatal(err)
	}

	// 遍歷所有列與欄
	for i, row := range rows {
		fmt.Printf("Row %d: ", i+1)
		for j, cell := range row {
			fmt.Printf("(%d,%d) %s  ", i+1, j+1, cell)
		}
		fmt.Println()
	}
}
